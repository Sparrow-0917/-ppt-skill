---
name: ppt-workflow
description: Reusable personal workflow for creating, editing, restyling, and validating PowerPoint decks, especially academic, urban planning, architecture, GIS, internship, course, and conference presentations. Use for new PPT/PPTX creation, existing-deck edits, style changes, and slide quality checks.
metadata:
  short-description: Personal PowerPoint workflow and QA
---

# PPT Workflow

Use this skill as a lightweight router. Do not load every reference.

## Start here

1. Read [QUICK_START.md](QUICK_START.md).
2. Read [SKILL_INDEX.md](SKILL_INDEX.md).
3. Load only the references selected by the index.
4. Use the installed official `Presentations` skill for PPTX implementation details and validation.

For copy-ready invocation examples, read [COMMANDS.md](COMMANDS.md) only when the user asks how to call this skill or wants a reusable prompt.

## Non-negotiable safeguards

- Treat the user's current deck as authoritative.
- Back up an existing deck before any edit and write to a new output file.
- Do not change content, layout, slide order, images, data, fonts, or animation unless the user authorizes that category.
- Separate visual restyling from content revision.
- Render and inspect every final slide; automated validation alone is insufficient.
- Prefer the smallest set of references needed for the current task.

## Reference routing

- General layout and editable-output rules: [ppt_base.md](ppt_base.md)
- White-background minimal blue style: [ppt_minimal_blue.md](ppt_minimal_blue.md)
- Traditional Chinese blue palettes and sampling records: [ppt_color_palettes.md](ppt_color_palettes.md)
- Planning, architecture, GIS, and academic decks: [ppt_academic_planning.md](ppt_academic_planning.md)
- Safe edits to existing files: [ppt_existing_file_edit.md](ppt_existing_file_edit.md)
- Final review and regression checks: [ppt_quality_check.md](ppt_quality_check.md)
- Tool and implementation choices: [pptx_generation.md](pptx_generation.md)

Use [scripts/ppt_guard_snapshot.ps1](scripts/ppt_guard_snapshot.ps1) when a Windows PowerPoint edit must prove that text, geometry, pictures, slide count, and animation metadata stayed unchanged.
