use super::super::super::class_exports::DotnetClassExport;
use super::component_classes_common::constructor_class;

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        constructor_class("dotnet.System.Data", "DataTable", "vybe:data", "dataTableNew"),
        constructor_class("dotnet.System.Data", "DataSet", "vybe:data", "dataSetNew"),
        constructor_class("dotnet.System.Drawing", "Point", "vybe:drawing", "pointNew"),
        constructor_class("dotnet.System.Drawing", "Size", "vybe:drawing", "sizeNew"),
        constructor_class("dotnet.System.Drawing", "SizeF", "vybe:drawing", "sizeNew"),
        constructor_class("dotnet.System.Drawing", "Font", "vybe:drawing", "fontNew"),
        constructor_class("dotnet.System.Drawing", "Pen", "vybe:drawing", "penNew"),
        constructor_class("dotnet.System.Drawing", "SolidBrush", "vybe:drawing", "solidBrushNew"),
        constructor_class("dotnet.System.Drawing", "Color", "vybe:drawing", "colorFromName"),
        constructor_class("dotnet.System.Drawing", "Graphics", "vybe:drawing", "graphicsNew"),
        constructor_class("dotnet.System.Data.SqlClient", "SqlConnection", "vybe:database", "connect"),
    ]
}
