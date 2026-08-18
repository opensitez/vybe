//! `DataGridView` — the WinForms data grid.
//!
//! Inherits from `Control` directly. Property surface is large; we list
//! the user-facing ones (cells, columns, rows, behaviour, appearance).

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[DotnetClass {
        name: "DataGridView",
        parent: Some("Control"),
        properties: &[
            "AllowUserToAddRows",
            "AllowUserToDeleteRows",
            "AllowUserToOrderColumns",
            "AllowUserToResizeColumns",
            "AllowUserToResizeRows",
            "AlternatingRowsDefaultCellStyle",
            "AutoGenerateColumns",
            "AutoSizeColumnsMode",
            "AutoSizeRowsMode",
            "BackgroundColor",
            "BorderStyle",
            "CellBorderStyle",
            "ClipboardCopyMode",
            "ColumnCount",
            "ColumnHeadersBorderStyle",
            "ColumnHeadersDefaultCellStyle",
            "ColumnHeadersHeight",
            "ColumnHeadersHeightSizeMode",
            "ColumnHeadersVisible",
            "Columns",
            "CurrentCell",
            "CurrentCellAddress",
            "CurrentRow",
            "DataMember",
            "DataSource",
            "DefaultCellStyle",
            "EditMode",
            "EnableHeadersVisualStyles",
            "FirstDisplayedCell",
            "FirstDisplayedScrollingColumnIndex",
            "FirstDisplayedScrollingRowIndex",
            "GridColor",
            "MultiSelect",
            "NewRowIndex",
            "ReadOnly",
            "RowCount",
            "RowHeadersBorderStyle",
            "RowHeadersDefaultCellStyle",
            "RowHeadersVisible",
            "RowHeadersWidth",
            "RowHeadersWidthSizeMode",
            "Rows",
            "RowsDefaultCellStyle",
            "RowTemplate",
            "ScrollBars",
            "SelectedCells",
            "SelectedColumns",
            "SelectedRows",
            "SelectionMode",
            "ShowCellErrors",
            "ShowCellToolTips",
            "ShowEditingIcon",
            "ShowRowErrors",
            "Sort",
            "SortedColumn",
            "SortOrder",
            "StandardTab",
            "TopLeftHeaderCell",
            "VirtualMode",
        ],
        methods: &[],
        ctor_arity: 0,
        widget_host_fn: None,    },
    // The legacy `DataGrid`, still nameable in an old `.Designer.vb`.
    // `control_kind` already folds `datagrid` onto the `datagridview` widget,
    // so declaring it costs one entry and it renders as the grid it is.
    DotnetClass {
        name: "DataGrid",
        parent: Some("Control"),
        properties: &[
            "AllowSorting",
            "CaptionText",
            "CaptionVisible",
            "CurrentCell",
            "DataMember",
            "DataSource",
            "ReadOnly",
        ],
        methods: &[],
        ctor_arity: 0,
        widget_host_fn: None,    },
    // ⚠ `PropertyGrid` has no widget kind yet, so it renders as a LABEL until
    // `vybe_widgets` grows one. That is the designed degradation for a
    // `vybe-*` tag naming a control the widget layer does not know — visible
    // in a capture and in `html`, rather than the control vanishing. The
    // DECLARATION is still worth having on its own: it makes the class
    // constructible, gives it identity, and makes its geometry, text and
    // events reach the document like any other control.
    DotnetClass {
        name: "PropertyGrid",
        parent: Some("Control"),
        properties: &[
            "CommandsVisibleIfAvailable",
            "HelpVisible",
            "PropertySort",
            "SelectedObject",
            "ToolbarVisible",
        ],
        methods: &[],
        ctor_arity: 0,
        widget_host_fn: None,    }]
}
