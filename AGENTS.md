# Repository Guidelines

## Purpose
This repository contains a small Windows command-line utility for inspecting and removing Explorer context-menu entries. The tool is expected to work directly against the Windows registry and lean on the Win32 API.

## Implementation Style
- Keep the code flat and procedural.
- Prefer C-style data structures, free functions, fixed buffers, and explicit control flow.
- Do not introduce class hierarchies, abstract interfaces, templates, smart-pointer frameworks, or layered architecture unless the user explicitly asks for them.
- Favor direct Win32 API calls over helper libraries.
- Keep dependencies at zero beyond the Windows SDK and the compiler runtime.

## Build Expectations
- The project builds with a single `build.bat`.
- `build.bat` should compile every `*.cpp` file under `src`.
- The produced executable belongs in `package`.
- Keep the build script easy to read and easy to run from a normal shell.

## CLI Expectations
- The program is a console application.
- It should expose simple verbs such as `find`, `remove`, `scopes`, and `help`.
- Output should be plain text and useful in a terminal without extra formatting layers.
- Registry operations should surface Windows error details when something fails.

## Safety
- Removing context-menu entries changes the registry. Keep the behavior explicit and predictable.
- Avoid broad or fuzzy deletion logic. Prefer exact scope and exact entry matches.
- If elevation is required, say so plainly.

## Change Discipline
- Preserve the flat style of the codebase.
- Do not add unrelated files, frameworks, or build systems.
- Validate changes by running `build.bat` when the local environment allows it.
