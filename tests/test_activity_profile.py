"""Tests for activity profile suggestions. Activity names are generic examples."""

import pytest

from wayfinder.activity_profile import ENVIRONMENTS, INTENSITIES, suggest_profile


@pytest.mark.parametrize("name, environment, intensity", [
    ("Garden Club", "outdoor", "light_active"),
    ("Morning Yoga", "unknown", "light_active"),
    ("Cooking Class", "indoor", "light_active"),
    ("Craft Corner", "indoor", "sitting_or_low"),
    ("Pickup Basketball", "unknown", "heavy_active"),
    ("Rake the Lawn", "outdoor", "moderate_active"),
    ("Walking Track", "unknown", "light_active"),
    ("Neighbourhood Walk", "outdoor", "light_active"),
    ("Visit Day", "unknown", "unknown"),
])
def test_suggestions(name, environment, intensity):
    profile = suggest_profile(name)
    assert profile["environment"]["suggested"] == environment
    assert profile["intensity"]["suggested"] == intensity


def test_every_suggestion_explains_itself_and_awaits_confirmation():
    profile = suggest_profile("Neighbourhood Walk")
    for dimension in ("environment", "intensity"):
        assert profile[dimension]["reason"]
        assert profile[dimension]["confirmed"] is None
    assert profile["environment"]["confidence"] == "medium"
    assert profile["review_status"] == "pending"


def test_unknown_has_no_confidence_level():
    profile = suggest_profile("Visit Day")
    assert profile["environment"]["confidence"] is None
    assert profile["intensity"]["confidence"] is None


def test_suggestions_stay_inside_the_agreed_categories():
    for name in ["Swim", "Bowling", "Park Clean-up", "Movie", "Dance", "Shopping Trip", "Unknown thing"]:
        profile = suggest_profile(name)
        assert profile["environment"]["suggested"] in ENVIRONMENTS
        assert profile["intensity"]["suggested"] in INTENSITIES


def test_reason_quotes_the_whole_word_from_the_schedule():
    assert '"Groceries"' in suggest_profile("Weekly Groceries")["environment"]["reason"]
