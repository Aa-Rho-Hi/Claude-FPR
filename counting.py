"""
counting.py — Robust, auditable entry-counting for CV / FAR sections.

Design principle (the whole point of this module)
--------------------------------------------------
Counting, segmentation, and filtering are SEPARATE concerns:

  1. extract_section()  — slice the right section out of the document.
  2. segment_entries()  — split that section into a LIST of discrete entries.
  3. a filter          — decide which entries match the rule (keyword OR an AI
                          per-entry yes/no classifier).
  4. the count          — ALWAYS `len(matched)`, computed in Python.

Neither regex NOR the AI ever *emits a number*. The AI is only ever asked
"does THIS one entry match? yes/no", one entry at a time, and Python tallies
the yeses. This structurally eliminates the class of bug where an LLM miscounts
a long list (13, then 36, then 55...) because nothing ever holds a running
tally over a large list.

Every result is auditable: run_rule() returns not just the number but the exact
entries that were counted, the segmentation method used, and a confidence level,
so any count can be verified at a glance on unseen data.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
from dataclasses import dataclass, field
from typing import Callable, List, Optional


# ══════════════════════════════════════════════════════════════════════════════
# Result types
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class Judgment:
    """One AI verdict on one entry."""
    match: bool
    reason: str = ""
    uncertain: bool = False   # True when self-consistency votes disagreed


@dataclass
class RuleResult:
    count: int
    matched: List[str] = field(default_factory=list)   # entries that were counted
    entries: List[str] = field(default_factory=list)   # ALL segmented entries
    method: str = ""          # how the section was segmented
    confidence: str = "high"  # high | medium | low
    mode: str = "regex"       # regex | ai
    warnings: List[str] = field(default_factory=list)
    reasons: List[tuple] = field(default_factory=list)   # (entry, reason) for counted entries
    uncertain: List[str] = field(default_factory=list)   # entries the AI was unsure about

    def as_int(self) -> int:
        return self.count


# ══════════════════════════════════════════════════════════════════════════════
# 1. Section extraction
# ══════════════════════════════════════════════════════════════════════════════

_ENTRY_MARKER = re.compile(r'^\s*(?:\d+[.)]|[\[(]\s*[A-Za-z]{0,3}\d+\s*[\])]|[•·▪‣◦∙*\-–—])\s')


def _norm(s: str) -> str:
    """Normalise for header matching: lowercase, collapse ws, strip punctuation."""
    s = re.sub(r'[^a-z0-9 ]', ' ', s.lower())
    return re.sub(r'\s+', ' ', s).strip()


def _is_header_line(line: str, next_line: str = "") -> bool:
    """
    Generic, format-agnostic section-header detector. A header is a short,
    title-ish line that does NOT start like an entry. We deliberately accept
    ALL-CAPS, Title Case, trailing-colon, and underlined headers — the old
    code only recognised ALL-CAPS >=6 chars, which is why sections bled.
    """
    s = line.strip()
    if not s or len(s) > 80:
        return False
    if _ENTRY_MARKER.match(s):
        return False
    words = s.split()
    if len(words) > 8:
        return False
    core = s.rstrip(':').strip()
    if not core:
        return False

    all_caps = core.upper() == core and any(c.isalpha() for c in core)
    title_case = bool(re.match(r'^([A-Z][\w&/\-]*)(\s+([A-Za-z&/\-]+))*$', core)) and \
        sum(1 for w in words if w[:1].isupper()) >= max(1, len(words) - 1)
    ends_colon = s.endswith(':')
    underlined = bool(re.match(r'^[=\-_~]{3,}$', next_line.strip()))

    return all_caps or title_case or ends_colon or underlined


def extract_section(cv_text: str, section: str) -> tuple[str, List[str]]:
    """
    Return (section_body_text, warnings).

    Strategy: find every header line in the document, pick the one that best
    matches `section`, then take everything from after that header up to the
    next header. If no section is requested, the whole document is the body.
    Falls back to a guarded substring match if no header matches.
    """
    warnings: List[str] = []
    if not section or not section.strip():
        return cv_text, warnings

    target = _norm(section)
    lines = cv_text.split("\n")

    # Locate all header lines.
    header_idx = [
        i for i, ln in enumerate(lines)
        if _is_header_line(ln, lines[i + 1] if i + 1 < len(lines) else "")
    ]

    # Score header matches against the target.
    best_i, best_score = None, 0
    for i in header_idx:
        h = _norm(lines[i])
        if not h:
            continue
        if h == target:
            score = 4
        elif h.startswith(target) or target.startswith(h):
            score = 3
        elif re.search(r'\b' + re.escape(target) + r'\b', h) or \
                re.search(r'\b' + re.escape(h) + r'\b', target):
            score = 2
        elif target in h:
            score = 1
        else:
            score = 0
        if score > best_score:
            best_score, best_i = score, i

    if best_i is not None:
        # Body runs until the next header line.
        nxt = next((j for j in header_idx if j > best_i), len(lines))
        body = "\n".join(lines[best_i + 1:nxt])
        if not body.strip():
            warnings.append(f"Section '{section}' matched a header but the body was empty.")
        return body, warnings

    # Fallback: substring, but only on short (header-like) lines to avoid
    # matching the word inside an entry body.
    for i, ln in enumerate(lines):
        if target in _norm(ln) and len(ln.split()) <= 8:
            nxt = next((j for j in header_idx if j > i), len(lines))
            warnings.append(f"Section '{section}' not found as a clean header; "
                            f"used a relaxed match on line {i + 1}.")
            return "\n".join(lines[i + 1:nxt]), warnings

    warnings.append(f"Section '{section}' not found — counted across the whole document.")
    return cv_text, warnings


# ══════════════════════════════════════════════════════════════════════════════
# 2. Entry segmentation (the detector cascade)
# ══════════════════════════════════════════════════════════════════════════════

_NUMBERED = re.compile(r'^\s*(\d+)[.)]\s+\S')
_BRACKETED = re.compile(r'^\s*[\[(]\s*([A-Za-z]{0,3}\d+)\s*[\])]\s*\S')
_BULLET = re.compile(r'^\s*([•·▪‣◦∙*]|[\-–—])\s+\S')


def _group_by_marker(lines: List[str], marker: re.Pattern) -> List[str]:
    """
    Group lines into entries: a line that matches `marker` STARTS a new entry;
    everything after it (until the next marker) is a continuation of that entry.
    This is what fixes multi-line citations being counted as multiple entries.
    """
    entries: List[str] = []
    cur: List[str] = []
    for ln in lines:
        if marker.match(ln):
            if cur:
                entries.append(" ".join(cur).strip())
            cur = [ln.strip()]
        elif cur:
            cur.append(ln.strip())
    if cur:
        entries.append(" ".join(cur).strip())
    return [e for e in entries if e]


def _numbered_is_sane(lines: List[str]) -> bool:
    """Reject false positives (e.g. years) by checking the numbers look like a
    list: they should start low and be mostly non-decreasing."""
    nums = [int(m.group(1)) for ln in lines if (m := _NUMBERED.match(ln))]
    if len(nums) < 2:
        return len(nums) == 1
    if nums[0] > 5:                       # real lists almost always start at 1..5
        return False
    inversions = sum(1 for a, b in zip(nums, nums[1:]) if b < a)
    return inversions <= max(1, len(nums) // 10)


def segment_entries(section_text: str) -> tuple[List[str], str, str]:
    """
    Return (entries, method, confidence).

    Detectors are tried by how explicit/reliable the structure is. We pick the
    DOMINANT explicit marker present in the section rather than the first one we
    see, then fall back to blank-line blocks, then (flagged) one-per-line.
    """
    raw_lines = section_text.split("\n")
    lines = [ln for ln in raw_lines if ln.strip()]
    if not lines:
        return [], "empty", "high"

    n = len(lines)
    num_hits = sum(1 for ln in lines if _NUMBERED.match(ln))
    brk_hits = sum(1 for ln in lines if _BRACKETED.match(ln))
    bul_hits = sum(1 for ln in lines if _BULLET.match(ln))

    # Explicit markers: use whichever clearly dominates the section start-lines.
    if num_hits >= 2 and num_hits >= brk_hits and num_hits >= bul_hits and _numbered_is_sane(lines):
        return _group_by_marker(lines, _NUMBERED), "numbered", "high"
    if brk_hits >= 2 and brk_hits >= bul_hits:
        return _group_by_marker(lines, _BRACKETED), "bracketed", "high"
    if bul_hits >= 2:
        return _group_by_marker(lines, _BULLET), "bulleted", "high"
    # A single explicit marker still counts as one structured list.
    if num_hits == 1 and n <= 2 and _numbered_is_sane(lines):
        return _group_by_marker(lines, _NUMBERED), "numbered", "high"

    # Blank-line separated paragraphs (common in narrative CVs).
    blocks = [b.strip() for b in re.split(r'\n\s*\n', section_text) if b.strip()]
    if len(blocks) >= 2:
        # Confidence drops if blocks are wildly uneven (could be prose, not a list).
        avg = sum(len(b) for b in blocks) / len(blocks)
        conf = "high" if avg < 600 else "medium"
        return [re.sub(r'\s+', ' ', b) for b in blocks], "blank-line blocks", conf

    # Single block but multiple sentences? Could be inline prose — flag it.
    if len(blocks) == 1 and n > 1:
        return [re.sub(r'\s+', ' ', b) for b in blocks], "single block", "low"

    # Last resort: one entry per non-blank line. LOW confidence, always flagged,
    # because multi-line entries with no markers will be over-counted here.
    return [ln.strip() for ln in lines], "one-per-line (unstructured)", "low"


# ══════════════════════════════════════════════════════════════════════════════
# 3. Filters
# ══════════════════════════════════════════════════════════════════════════════

def _keyword_filter(entries: List[str], rule_type: str, keywords: List[str],
                    year: str) -> List[str]:
    kws = [k.lower() for k in keywords if k]
    out = []
    for e in entries:
        el = e.lower()
        if rule_type == "all":
            ok = True
        elif rule_type == "contains":
            ok = bool(kws) and kws[0] in el
        elif rule_type == "year":
            ok = bool(year) and year in el
        elif rule_type == "all_of":
            ok = bool(kws) and all(k in el for k in kws)
        elif rule_type == "any_of":
            ok = bool(kws) and any(k in el for k in kws)
        elif rule_type == "excludes":
            ok = bool(kws) and kws[0] not in el
        else:
            ok = False
        if ok:
            out.append(e)
    return out


# ══════════════════════════════════════════════════════════════════════════════
# 4. Orchestrator
# ══════════════════════════════════════════════════════════════════════════════

def run_rule(cv_text: str, rule: dict,
             ai_classifier: Optional[Callable[[List[str], str], List[bool]]] = None
             ) -> RuleResult:
    """
    Execute one rule against cv_text.

    rule keys:
      section    : section heading to look in ("" = whole document)
      mode        : "regex" (default) or "ai"
      rule_type   : all|contains|year|all_of|any_of|excludes  (regex mode)
      keywords    : list[str]                                   (regex mode)
      year        : str                                         (regex mode)
      instruction : natural-language test                       (ai mode)

    ai_classifier(entries, instruction) -> list[bool] of the same length.
    If None and mode == "ai", we degrade gracefully (count 0 + warning).
    """
    section = rule.get("section", "")
    body, warnings = extract_section(cv_text, section)
    entries, method, confidence = segment_entries(body)

    mode = rule.get("mode", "regex")

    if mode == "ai":
        instruction = rule.get("instruction", "").strip()
        if not entries:
            return RuleResult(0, [], entries, method, confidence, "ai",
                              warnings + ["No entries found to classify."])
        if ai_classifier is None:
            return RuleResult(
                0, [], entries, method, confidence, "ai",
                warnings + ["AI rule could not run: no API key/classifier configured. "
                            "Set ANTHROPIC_API_KEY (or OPENAI_API_KEY)."])
        try:
            verdicts = ai_classifier(entries, instruction)
        except Exception as e:  # pragma: no cover - network/SDK errors
            return RuleResult(0, [], entries, method, confidence, "ai",
                              warnings + [f"AI classification failed: {e}"])

        # Accept either rich Judgments or bare booleans (test mocks / simple callers).
        judged = [j if isinstance(j, Judgment) else Judgment(bool(j)) for j in verdicts]
        matched   = [e for e, j in zip(entries, judged) if j.match]
        reasons   = [(e, j.reason) for e, j in zip(entries, judged) if j.match]
        uncertain = [e for e, j in zip(entries, judged) if j.uncertain]
        w = list(warnings)
        if uncertain:
            w.append(f"AI was unsure about {len(uncertain)} entr"
                     f"{'y' if len(uncertain) == 1 else 'ies'} (votes disagreed) — review these.")
        return RuleResult(len(matched), matched, entries, method, confidence, "ai",
                          w, reasons=reasons, uncertain=uncertain)

    # regex mode
    matched = _keyword_filter(
        entries,
        rule.get("rule_type", "contains"),
        rule.get("keywords", []) or [],
        str(rule.get("year", "")),
    )
    return RuleResult(len(matched), matched, entries, method, confidence, "regex", warnings)


# ══════════════════════════════════════════════════════════════════════════════
# 5. AI per-entry classifier (Anthropic default, OpenAI fallback)
# ══════════════════════════════════════════════════════════════════════════════

_AI_CACHE: dict[str, Judgment] = {}


def _cache_key(instruction: str, entry: str) -> str:
    return hashlib.sha256((instruction + "\x00" + entry).encode("utf-8")).hexdigest()


def make_ai_classifier(api_key: Optional[str] = None,
                       provider: str = "auto",
                       model: Optional[str] = None,
                       batch_size: int = 20,
                       votes: int = 3) -> Optional[Callable[[List[str], str], List[Judgment]]]:
    """
    Build a classifier(entries, instruction) -> list[Judgment].

    The model is shown a NUMBERED batch of entries and asked to return, for each
    id, {match: true/false, reason: "..."} against the instruction. It never
    returns a total — Python counts the matches.

    Self-consistency: each batch is judged `votes` times at temperature 0 and the
    per-entry verdict is the MAJORITY. If the votes for an entry are not unanimous
    it is marked `uncertain` so a human can review it. Results are cached per
    (instruction, entry) so reruns are stable and cheap. Returns None if no
    provider/key is available (caller degrades gracefully).
    """
    key = api_key or os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("OPENAI_API_KEY")
    if not key:
        return None

    use = provider
    if provider == "auto":
        if os.environ.get("ANTHROPIC_API_KEY") or api_key:
            use = "anthropic"
        elif os.environ.get("OPENAI_API_KEY"):
            use = "openai"
        else:
            use = "anthropic"

    def _call_model(prompt: str) -> str:
        if use == "anthropic":
            import anthropic
            client = anthropic.Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))
            msg = client.messages.create(
                model=model or "claude-sonnet-4-5",
                max_tokens=1024,
                messages=[{"role": "user", "content": prompt}],
            )
            return "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
        else:
            from openai import OpenAI
            client = OpenAI(api_key=api_key or os.environ.get("OPENAI_API_KEY"))
            resp = client.chat.completions.create(
                model=model or "gpt-4o-mini",
                temperature=0,
                messages=[{"role": "user", "content": prompt}],
            )
            return resp.choices[0].message.content or ""

    def _parse_one_vote(text: str, n: int) -> tuple[List[bool], List[str]]:
        """Parse one model response into (match flags, reasons) of length n."""
        m = re.search(r'\[.*\]', text, re.DOTALL)
        if not m:
            raise ValueError(f"model did not return JSON: {text[:200]}")
        data = json.loads(m.group(0))
        flags = [False] * n
        reasons = [""] * n
        for item in data:
            i = int(item.get("id"))
            if 0 <= i < n:
                flags[i] = bool(item.get("match"))
                reasons[i] = str(item.get("reason", ""))[:200]
        return flags, reasons

    def _judge_batch(entries: List[str], idxs: List[int],
                     instruction: str) -> List[Judgment]:
        """Judge one batch `votes` times and take the per-entry majority."""
        listing = "\n".join(
            f"[{j}] {re.sub(chr(10), ' ', entries[g])[:500]}" for j, g in enumerate(idxs)
        )
        prompt = (
            "You are labelling CV entries. For EACH entry below, decide whether it "
            f"satisfies this condition:\n\n  CONDITION: {instruction}\n\n"
            "Judge each entry independently and literally. Do NOT count or summarise. "
            "Return ONLY a JSON array, one object per entry, like "
            '[{"id":0,"match":true,"reason":"filed 2018, after 2015"},'
            '{"id":1,"match":false,"reason":"filed 2012"}]. '
            "Keep each reason under 15 words.\n\n"
            f"ENTRIES:\n{listing}"
        )
        tally = [0] * len(idxs)          # how many votes said True
        first_reasons = [""] * len(idxs)
        for v in range(max(1, votes)):
            flags, reasons = _parse_one_vote(_call_model(prompt), len(idxs))
            for j in range(len(idxs)):
                tally[j] += int(flags[j])
                if v == 0:
                    first_reasons[j] = reasons[j]
        out = []
        for j in range(len(idxs)):
            n_votes = max(1, votes)
            match = tally[j] * 2 > n_votes          # strict majority
            unanimous = tally[j] == 0 or tally[j] == n_votes
            out.append(Judgment(match=match, reason=first_reasons[j],
                                uncertain=not unanimous))
        return out

    def classifier(entries: List[str], instruction: str) -> List[Judgment]:
        results: List[Optional[Judgment]] = [None] * len(entries)

        # Serve from cache; collect the misses.
        todo: List[int] = []
        for i, e in enumerate(entries):
            ck = _cache_key(instruction, e)
            if ck in _AI_CACHE:
                results[i] = _AI_CACHE[ck]
            else:
                todo.append(i)

        for start in range(0, len(todo), batch_size):
            idxs = todo[start:start + batch_size]
            verdicts = _judge_batch(entries, idxs, instruction)
            for j, g in enumerate(idxs):
                results[g] = verdicts[j]
                _AI_CACHE[_cache_key(instruction, entries[g])] = verdicts[j]

        return [r if isinstance(r, Judgment) else Judgment(False) for r in results]

    return classifier
