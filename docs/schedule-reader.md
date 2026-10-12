# Schedule Reader v0

Schedule Reader turns a weekly schedule written in Word (.docx) into structured
JSON that a guardian or supporter can check, and that later Wayfinder features
can use.

It works with the schedule people already write. The person who writes the
schedule does not need to change how they write it.

## What the schedule should look like

The reader is built for a common hand-made layout:

- a few lines of text above a table, such as a title and a theme;
- one table whose first row has day headings such as `Mon Mar 1`;
- one row per part of the day, labelled `MORNING`, `AFTERNOON`, and so on;
- in each cell, numbered activities (`1. Yoga`), each followed by a plain
  (not bold) line listing who is assigned to it.

Anything else in a cell, such as `Closed`, a holiday greeting or a menu, is
kept word for word as a note.

## Three kinds of information

| Kind | Where it appears | Who decides it |
|------|------------------|----------------|
| **Source facts** | `title`, `heading`, `label`, `name`, `source_text`, `participants`, `notes` | The schedule document |
| **Wayfinder suggestions** | `classification.*.suggested`, `confidence`, `reason`, plus `date` and `date_range` | Wayfinder (always marked or explained) |
| **Human confirmation** | `classification.*.confirmed`, `review_status`, `corrections` | A guardian or supporter |

The reader always leaves `confirmed` as `null` and `review_status` as
`pending`.

## Activity profile

Each activity gets two separate suggestions:

- **environment**: `indoor`, `outdoor`, `mixed` or `unknown`
- **intensity**: `sitting_or_low`, `light_active`, `moderate_active`,
  `heavy_active` or `unknown`

Suggestions come from simple word rules in `wayfinder/activity_profile.py`,
and each one includes the reason it was made. If the words give no reasonable
clue, the suggestion is `unknown` and its `confidence` is `null`.

## Honesty about uncertainty

Wayfinder does not guess. Instead it adds an item to `needs_review` when, for
example:

- a part of the day has text but no activities (it is unclear whether the
  program runs);
- an activity has no list of participants;
- a line could belong to one activity or to the whole period;
- a name entry uses brackets or hyphens that the schedule does not explain;
- a day heading has no year, or its weekday does not match its date (the
  `date` is then left as `null`);
- an activity's environment or intensity suggestion is anything less than
  high confidence (medium, low or unknown). The item names which suggestion
  to check and how confident Wayfinder was.

## How a guardian confirms or corrects

The guardian screen has not been built yet. The format is ready for it:

- **To confirm or change a suggestion**, set `confirmed` to one of the
  categories (including `unknown`) and set `review_status` to `confirmed` or
  `changed`. `suggested` is never overwritten.
- **To correct something the reader extracted**, add an entry to
  `corrections` rather than editing the source text:

```json
{
  "where": "day2-morning-1",
  "field": "name",
  "source_value": "Sprng Hike",
  "corrected_value": "Spring Hike",
  "note": "Typo in the schedule"
}
```

Every `where` value is an `id` that appears in the output.

## Running it

```
pip install -e .[test]
python -m wayfinder.schedule_reader "path/to/schedule.docx" -o "path/outside/the/repo/schedule.json"
pytest
```

Real schedules contain personal information. Real schedules, and the JSON made
from them, belong **outside** this repository entirely. The reader refuses to
write output inside the Wayfinder folder, and `.gitignore` excludes `.docx`
files as a second safeguard.

The automated tests use invented schedules that are built in code while the
tests run, so no Word files are stored in the repository.
