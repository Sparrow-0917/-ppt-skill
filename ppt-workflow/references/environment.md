# Current PPT Environment

## Available capabilities

- Official Codex `Presentations` skill with rendering and finalization tools.
- `@oai/artifact-tool` runtime for editable PPTX authoring and inspection.
- Microsoft PowerPoint COM automation on Windows.
- Bundled Python runtime and presentation package validators.
- LibreOffice-based rendering fallback.

## Decision rule

Use the official Presentations workflow by default. Use PowerPoint COM only when an existing complex file must retain native PowerPoint features during narrowly scoped edits. Validate and render after either path.

## External resources

No third-party repository is currently necessary. Review this decision only when a task requires a capability not available above.
