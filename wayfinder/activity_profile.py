"""Suggest an activity profile (environment and intensity) from an activity name.

These are SUGGESTIONS made by Wayfinder, never facts from the schedule.
They are made with simple, readable word rules so that anyone can see why a
suggestion was made. When the words give no reasonable clue, the answer is
"unknown".

Each dimension is suggested independently, so an activity can have a known
intensity and an unknown environment (for example "Basketball": clearly
active, but it could be in a gym or on an outdoor court).
"""

import re

ENVIRONMENTS = ("indoor", "outdoor", "mixed", "unknown")
INTENSITIES = ("sitting_or_low", "light_active", "moderate_active", "heavy_active", "unknown")

# Rules are checked in order; the first matching rule wins.
# Each rule: (regular expression, suggested value, confidence, explanation).
ENVIRONMENT_RULES = [
    (r"\btrack\b", "unknown", None,
     "a track can be indoors or outdoors and the schedule does not say which"),
    (r"basketball|volleyball", "unknown", None,
     "this can be played in a gym or on an outdoor court and the schedule does not say which"),
    (r"\bswim", "indoor", "low",
     "swimming is often at an indoor pool, but the schedule does not say so"),
    (r"\bgym\b", "indoor", "medium", "a gym is normally indoors"),
    (r"rak(e|ing)|\bleaf\b|\bleaves\b|garden|hik(e|ing)|picnic", "outdoor", "high",
     "this is normally done outdoors"),
    (r"harvest", "outdoor", "medium",
     "harvesting usually happens outdoors, but the schedule does not say where"),
    (r"\bpark\b", "outdoor", "medium",
     "a park suggests outdoors, but it may only be the name of the place and part of the time could be inside"),
    (r"kitchen|cook|bak(e|ing)|dish|crafts?\b|\bart\b|movie|library|bowling|board games?", "indoor", "high",
     "this is normally done indoors"),
    (r"grocer|shopping|\bstore\b|\bmall\b", "indoor", "medium",
     "shopping happens inside a store, but getting there means travelling outside"),
    (r"\bwalk", "outdoor", "medium",
     "walks are usually outdoors, but the schedule does not say where"),
    (r"decorat", "indoor", "low",
     "decorating is often indoors, but it could be outdoors and the schedule does not say"),
]

INTENSITY_RULES = [
    (r"basketball|soccer|football|hockey|\brun(ning)?\b|zumba|aerobics", "heavy_active", "medium",
     "sports and running are usually vigorous, though how hard people play varies"),
    (r"obstacle|rak(e|ing)|shovel|danc(e|ing)|hik(e|ing)|swim|bik(e|ing)|cycling", "moderate_active", "medium",
     "this usually involves steady physical effort"),
    (r"yoga|stretch", "light_active", "high", "this is gentle movement"),
    (r"\bwalk", "light_active", "high", "walking is light activity"),
    (r"garden|harvest|scavenger|grocer|shopping|kitchen|cook|bak(e|ing)|dish|clean|decorat|bowling",
     "light_active", "medium", "this usually means standing and moving around without much exertion"),
    (r"crafts?\b|\bart\b|movie|library|board games?|reading|music|bingo", "sitting_or_low", "medium",
     "this is usually done sitting or with little movement"),
]

NO_CLUE = {
    "environment": "The activity name does not say whether it is indoors or outdoors.",
    "intensity": "The activity name does not say how physically active it is.",
}


def _whole_word(text, match):
    """The full word(s) around a match, so reasons quote the schedule's own words."""
    start, end = match.start(), match.end()
    while start > 0 and text[start - 1].isalnum():
        start -= 1
    while end < len(text) and text[end].isalnum():
        end += 1
    return text[start:end]


def _suggest(name, rules, dimension):
    for pattern, value, confidence, explanation in rules:
        match = re.search(pattern, name, re.IGNORECASE)
        if match:
            reason = f'The activity name contains "{_whole_word(name, match)}": {explanation}.'
            return {"suggested": value, "confidence": confidence if value != "unknown" else None,
                    "reason": reason, "confirmed": None}
    return {"suggested": "unknown", "confidence": None, "reason": NO_CLUE[dimension], "confirmed": None}


def suggest_profile(activity_name):
    """Return Wayfinder's suggested profile for one activity name.

    "confirmed" is always None here: only a guardian or supporter fills it in.
    """
    return {
        "environment": _suggest(activity_name, ENVIRONMENT_RULES, "environment"),
        "intensity": _suggest(activity_name, INTENSITY_RULES, "intensity"),
        "review_status": "pending",
    }
