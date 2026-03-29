# vybe_editor → vybe_editor2 Migration Plan

Porting the missing features from the original Dioxus-based `vybe_editor` into the egui-based `vybe_editor2`.

## Features

| # | Feature | File | Status |
|---|---------|------|--------|
| 1 | **Data section in Toolbox** (BindingSource, DataAdapter, DataSet, DataTable, BindingNavigator + dialogs/infra) | `panels/toolbox.rs` | ✅ Done |
| 2 | **Non-visual controls added immediately on toolbox click** (no canvas click needed) | `panels/toolbox.rs` | ✅ Done |
| 3 | **More standard controls** (MenuStrip, ContextMenuStrip, StatusStrip, ToolStrip, MaskedTextBox, SplitContainer, FlowLayoutPanel, TableLayoutPanel, MonthCalendar, HScrollBar, VScrollBar, WebBrowser, LinkLabel, RichTextBox) | `panels/toolbox.rs` | ✅ Done |
| 4 | **Component Tray** — non-visual components shown as icon chips below the form canvas | `panels/form_designer.rs` | ✅ Done |
| 5 | **Lasso / rubber-band selection** — drag on empty canvas to multi-select visual controls | `panels/form_designer.rs` | ✅ Done |
| 6 | **Multi-select drag** — all selected controls move together when any one is dragged | `panels/form_designer.rs` | ✅ Done |
| 7 | **Resize handles** — drag corners/edges to resize (single selection) | `panels/form_designer.rs` | ✅ Done |
| 8 | **Richer control visuals** — DataGridView headers, ListBox/ComboBox items, TreeView nodes, CheckBox/RadioButton widgets, ProgressBar fill, TrackBar | `panels/form_designer.rs` | ✅ Done |

## Files Changed

- `crates/vybe_editor2/src/panels/toolbox.rs` — rewritten with Controls + Data sections
- `crates/vybe_editor2/src/panels/form_designer.rs` — rewritten with all 5 canvas features

## Notes

- Non-visual controls bypass canvas placement and are added directly to the form on toolbox click.
- The component tray only renders when the form contains at least one non-visual control.
- Lasso state uses the existing `LassoState` struct in `state.rs` (origin + current egui::Pos2).
- Multi-select drag: on `drag_started`, snapshot all selected controls' bounds; on `dragged`, apply delta to all.
- Resize uses 8 handle positions (corners + mid-edges); only active for single-selection.
