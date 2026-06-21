# Progress - VBA Progress Manager

[![VBA](https://img.shields.io/badge/language-VBA-867db1.svg)](https://learn.microsoft.com/en-us/office/vba/api/overview/)
[![Version](https://img.shields.io/github/v/tag/naoto256/vba-progress?label=version&sort=semver)](https://github.com/naoto256/vba-progress/tags)
[![License: MIT OR Apache-2.0](https://img.shields.io/badge/license-MIT%20OR%20Apache--2.0-blue.svg)](#license)

Progress is a small VBA progress manager for composing nested progress reports.

Have you ever had to make every procedure know how much of the whole operation
it represents, just to show a progress bar?

That coupling spreads quickly: a leaf procedure that only imports rows starts
knowing that it is "35% to 80%" of the entire macro, and changing the workflow
means editing progress math throughout the call tree.

Progress solves that by making progress local. Callers decide where a child
operation fits; callees keep reporting their own `0..1` progress.

Features:

- normalize progress values to `0..1`
- compose nested progress scopes without leaking global percentages into child procedures
- notify one or more output listeners
- keep UI adapters separate from progress calculation

## Installation

Import the core class modules into your VBA project:

- `Progress.cls`
- `ProgressManager.cls`
- `ProgressScope.cls`
- `ProgressScopeGuard.cls`
- `IProgressListener.cls`

The `examples/ProgressListener_*.cls` files are sample listener adapters. Import
the ones you want, or implement `IProgressListener` in your own project:

- `examples/ProgressListener_DebugPrint.cls`
- `examples/ProgressListener_AppStatusBar.cls`
- `examples/ProgressListener_ProgressWindow.cls`

`examples/ProgressListener_ProgressWindow.cls` expects a project-specific
`ProgressWindow` form with compatible `OpenUI`, `SetProgressValue`, and
`CloseUI` members.

## Usage

```vb
Public Sub Example()
    Progress.AddListener "status", New ProgressListener_AppStatusBar
    Progress.AddListener "debug", New ProgressListener_DebugPrint

    Progress.Begin

    ' The top level owns only the global allocation.
    With Progress.Scope(0, 0.2)
        LoadInputFiles
    End With

    With Progress.Scope(0.2, 1)
        ImportWorkbook
    End With

    Progress.Finish
End Sub

Private Sub LoadInputFiles()
    Dim i As Long

    ' This procedure only knows about file loading. It reports local 0..1
    ' progress and does not know that the caller assigned it 20% globally.
    For i = 1 To 5
        ' Load one file...
        Progress.Update i / 5
    Next i
End Sub

Private Sub ImportWorkbook()
    Dim sheetIndex As Long

    ' This procedure owns its internal allocation. It still thinks in local
    ' 0..1 terms, even though the caller mapped it to global 0.2..1.
    For sheetIndex = 1 To 4
        With Progress.Scope((sheetIndex - 1) / 4, sheetIndex / 4)
            ImportSheet sheetIndex
        End With
    Next sheetIndex
End Sub

Private Sub ImportSheet(sheetIndex As Long)
    Dim rowIndex As Long

    ' This leaf procedure only knows about rows in one sheet.
    For rowIndex = 1 To 1000
        ' Import one row...
        Progress.Update rowIndex / 1000
    Next rowIndex
End Sub
```

`Progress.Scope(startPoint, endPoint)` maps local progress in the scope to the
given range in the parent scope. The called procedure can keep reporting local
`0..1` progress; only the caller decides where that work fits in the parent
operation.

Use it with `With ... End With` so the scope guard stays alive for the duration
of the block.

## Lifecycle

- `Progress.Begin` starts a new root operation and notifies listeners with `0`.
- `Progress.Update value` reports local progress in the active scope.
- `Progress.Finish` reports `1` and resets progress state.
- `Progress.Reset` resets progress state while keeping listeners.

`Begin`, `Reset`, and `Finish` must be called outside active scopes.

## Specs

The `specs/` directory contains executable VBA specs. Import the files under
`specs/` together with the core modules, then run:

```vb
RunAllProgressSpecs
```

The specs raise a VBA error on failure and print a success message on pass.

## License

Licensed under either of:

- Apache License, Version 2.0 (`LICENSE-APACHE`)
- MIT license (`LICENSE-MIT`)

at your option.
