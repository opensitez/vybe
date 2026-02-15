use crate::control::{Control, ControlType};
use crate::form::Form;
use crate::properties::PropertyValue;

/// Returns the field name for a control in VB.NET designer code.
/// Array members use `Name_Index` (e.g., `Command1_0`), non-array use `Name`.
fn control_field_name(control: &Control) -> String {
    if let Some(idx) = control.index {
        format!("{}_{}", control.name, idx)
    } else {
        control.name.clone()
    }
}

/// Maps a ControlType to its fully-qualified VB.NET class name.
pub fn control_type_to_vbnet(ct: &ControlType) -> &str {
    match ct {
        ControlType::Button => "System.Windows.Forms.Button",
        ControlType::Label => "System.Windows.Forms.Label",
        ControlType::TextBox => "System.Windows.Forms.TextBox",
        ControlType::CheckBox => "System.Windows.Forms.CheckBox",
        ControlType::RadioButton => "System.Windows.Forms.RadioButton",
        ControlType::ComboBox => "System.Windows.Forms.ComboBox",
        ControlType::ListBox => "System.Windows.Forms.ListBox",
        ControlType::Frame => "System.Windows.Forms.GroupBox",
        ControlType::PictureBox => "System.Windows.Forms.PictureBox",
        ControlType::RichTextBox => "System.Windows.Forms.RichTextBox",
        ControlType::WebBrowser => "System.Windows.Forms.WebBrowser",
        ControlType::TreeView => "System.Windows.Forms.TreeView",
        ControlType::DataGridView => "System.Windows.Forms.DataGridView",
        ControlType::Panel => "System.Windows.Forms.Panel",
        ControlType::ListView => "System.Windows.Forms.ListView",
        ControlType::BindingNavigator => "System.Windows.Forms.BindingNavigator",
        ControlType::TabControl => "System.Windows.Forms.TabControl",
        ControlType::TabPage => "System.Windows.Forms.TabPage",
        ControlType::ProgressBar => "System.Windows.Forms.ProgressBar",
        ControlType::NumericUpDown => "System.Windows.Forms.NumericUpDown",
        ControlType::MenuStrip => "System.Windows.Forms.MenuStrip",
        ControlType::ToolStripMenuItem => "System.Windows.Forms.ToolStripMenuItem",
        ControlType::ContextMenuStrip => "System.Windows.Forms.ContextMenuStrip",
        ControlType::StatusStrip => "System.Windows.Forms.StatusStrip",
        ControlType::ToolStripStatusLabel => "System.Windows.Forms.ToolStripStatusLabel",
        ControlType::DateTimePicker => "System.Windows.Forms.DateTimePicker",
        ControlType::LinkLabel => "System.Windows.Forms.LinkLabel",
        ControlType::ToolStrip => "System.Windows.Forms.ToolStrip",
        ControlType::TrackBar => "System.Windows.Forms.TrackBar",
        ControlType::MaskedTextBox => "System.Windows.Forms.MaskedTextBox",
        ControlType::SplitContainer => "System.Windows.Forms.SplitContainer",
        ControlType::FlowLayoutPanel => "System.Windows.Forms.FlowLayoutPanel",
        ControlType::TableLayoutPanel => "System.Windows.Forms.TableLayoutPanel",
        ControlType::MonthCalendar => "System.Windows.Forms.MonthCalendar",
        ControlType::HScrollBar => "System.Windows.Forms.HScrollBar",
        ControlType::VScrollBar => "System.Windows.Forms.VScrollBar",
        ControlType::ToolTip => "System.Windows.Forms.ToolTip",
        ControlType::BindingSourceComponent => "System.Windows.Forms.BindingSource",
        ControlType::DataSetComponent => "System.Data.DataSet",
        ControlType::DataTableComponent => "System.Data.DataTable",
        ControlType::DataAdapterComponent => "System.Data.SqlClient.SqlDataAdapter",
        ControlType::Timer => "System.Windows.Forms.Timer",
        ControlType::ImageList => "System.Windows.Forms.ImageList",
        ControlType::ErrorProvider => "System.Windows.Forms.ErrorProvider",
        ControlType::OpenFileDialog => "System.Windows.Forms.OpenFileDialog",
        ControlType::SaveFileDialog => "System.Windows.Forms.SaveFileDialog",
        ControlType::FolderBrowserDialog => "System.Windows.Forms.FolderBrowserDialog",
        ControlType::FontDialog => "System.Windows.Forms.FontDialog",
        ControlType::ColorDialog => "System.Windows.Forms.ColorDialog",
        ControlType::PrintDialog => "System.Windows.Forms.PrintDialog",
        ControlType::PrintDocument => "System.Drawing.Printing.PrintDocument",
        ControlType::NotifyIcon => "System.Windows.Forms.NotifyIcon",
        // Additional visual controls
        ControlType::CheckedListBox => "System.Windows.Forms.CheckedListBox",
        ControlType::DomainUpDown => "System.Windows.Forms.DomainUpDown",
        ControlType::PropertyGrid => "System.Windows.Forms.PropertyGrid",
        ControlType::Splitter => "System.Windows.Forms.Splitter",
        ControlType::DataGrid => "System.Windows.Forms.DataGrid",
        ControlType::UserControl => "System.Windows.Forms.UserControl",
        // ToolStrip sub-components
        ControlType::ToolStripSeparator => "System.Windows.Forms.ToolStripSeparator",
        ControlType::ToolStripButton => "System.Windows.Forms.ToolStripButton",
        ControlType::ToolStripLabel => "System.Windows.Forms.ToolStripLabel",
        ControlType::ToolStripComboBox => "System.Windows.Forms.ToolStripComboBox",
        ControlType::ToolStripDropDownButton => "System.Windows.Forms.ToolStripDropDownButton",
        ControlType::ToolStripSplitButton => "System.Windows.Forms.ToolStripSplitButton",
        ControlType::ToolStripTextBox => "System.Windows.Forms.ToolStripTextBox",
        ControlType::ToolStripProgressBar => "System.Windows.Forms.ToolStripProgressBar",
        // Additional dialogs
        ControlType::PrintPreviewDialog => "System.Windows.Forms.PrintPreviewDialog",
        ControlType::PageSetupDialog => "System.Windows.Forms.PageSetupDialog",
        ControlType::PrintPreviewControl => "System.Windows.Forms.PrintPreviewControl",
        // Non-visual infrastructure
        ControlType::HelpProvider => "System.Windows.Forms.HelpProvider",
        ControlType::BackgroundWorker => "System.ComponentModel.BackgroundWorker",
        ControlType::SqlConnection => "System.Data.SqlClient.SqlConnection",
        ControlType::OleDbConnection => "System.Data.OleDb.OleDbConnection",
        ControlType::DataView => "System.Data.DataView",
        ControlType::Custom(s) => s.as_str(),
    }
}

/// Format a PropertyValue as a VB.NET assignment RHS.
/// Strings are quoted with `""` escaping; booleans use True/False;
/// Expressions are emitted verbatim (raw VB.NET code).
fn property_value_to_vbnet(val: &PropertyValue) -> Option<String> {
    match val {
        PropertyValue::String(s) => Some(format!("\"{}\"", s.replace('"', "\"\""))),
        PropertyValue::Integer(i) => Some(i.to_string()),
        PropertyValue::Boolean(b) => Some(if *b { "True".to_string() } else { "False".to_string() }),
        PropertyValue::Double(d) => Some(d.to_string()),
        PropertyValue::Expression(code) => Some(code.clone()),
        // StringArray values need special handling (Items.AddRange etc.) – skip in generic output
        PropertyValue::StringArray(_) => None,
    }
}

/// Format a font string "Family, sizepx[, Style]" into a VB.NET Font constructor call.
fn format_font(font_str: &str) -> String {
    let mut parts = font_str.splitn(3, ',').map(|s| s.trim());
    let family = parts.next().unwrap_or("Segoe UI");
    let size_raw = parts.next().unwrap_or("12");
    let style_opt = parts.next(); // optional third part e.g. "System.Drawing.FontStyle.Bold"
    let size_clean = size_raw.trim_end_matches("px").trim_end_matches("pt");
    let size_val: f32 = size_clean.parse().unwrap_or(12.0);
    let family_escaped = family.replace('"', "\"\"");
    if let Some(style) = style_opt {
        format!(
            "New System.Drawing.Font(\"{}\", {}F, {})",
            family_escaped, size_val, style
        )
    } else {
        format!("New System.Drawing.Font(\"{}\", {}F)", family_escaped, size_val)
    }
}

/// Generates VB.NET designer code (InitializeComponent) from a Form object.
/// Output is compatible with real VB.NET Windows Forms designer files.
pub fn generate_designer_code(form: &Form) -> String {
    let mut code = String::new();

    code.push_str(&format!("Partial Class {}\n", form.name));
    code.push_str("    Inherits System.Windows.Forms.Form\n");
    code.push('\n');

    // Field declarations with Friend WithEvents (standard VB.NET pattern)
    for control in &form.controls {
        let vb_type = control_type_to_vbnet(&control.control_type);
        let field_name = control_field_name(control);
        code.push_str(&format!(
            "    Friend WithEvents {} As {}\n",
            field_name, vb_type
        ));
    }
    code.push('\n');

    code.push_str("    Private Sub InitializeComponent()\n");

    // ── 1. Instantiation ───────────────────────────────────────────────────
    for control in &form.controls {
        let vb_type = control_type_to_vbnet(&control.control_type);
        let field_name = control_field_name(control);
        code.push_str(&format!("        Me.{} = New {}()\n", field_name, vb_type));
    }
    code.push_str("        Me.SuspendLayout()\n");

    // ── 2. Per-control property assignments ────────────────────────────────
    for control in &form.controls {
        let field_name = control_field_name(control);
        let is_non_visual = control.control_type.is_non_visual();

        if !is_non_visual {
            code.push_str(&format!(
                "        Me.{}.Location = New System.Drawing.Point({}, {})\n",
                field_name, control.bounds.x, control.bounds.y
            ));
            code.push_str(&format!(
                "        Me.{}.Size = New System.Drawing.Size({}, {})\n",
                field_name, control.bounds.width, control.bounds.height
            ));
        }

        // Non-visual component specific properties
        if is_non_visual {
            // DataSource (reference to another component)
            if let Some(ds) = control.properties.get_string("DataSource") {
                if !ds.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.DataSource = Me.{}\n",
                        field_name, ds
                    ));
                }
            }
            if let Some(dm) = control.properties.get_string("DataMember") {
                if !dm.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.DataMember = \"{}\"\n",
                        field_name, dm
                    ));
                }
            }
            if let Some(filter) = control.properties.get_string("Filter") {
                if !filter.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.Filter = \"{}\"\n",
                        field_name, filter
                    ));
                }
            }
            if let Some(sort) = control.properties.get_string("Sort") {
                if !sort.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.Sort = \"{}\"\n",
                        field_name, sort
                    ));
                }
            }
            if let Some(tn) = control.properties.get_string("TableName") {
                if !tn.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.TableName = \"{}\"\n",
                        field_name, tn
                    ));
                }
            }
            if let Some(dsn) = control.properties.get_string("DataSetName") {
                if !dsn.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.DataSetName = \"{}\"\n",
                        field_name, dsn
                    ));
                }
            }
            if let Some(sc) = control.properties.get_string("SelectCommand") {
                if !sc.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.SelectCommand = \"{}\"\n",
                        field_name, sc
                    ));
                }
            }
            if let Some(cs) = control.properties.get_string("ConnectionString") {
                if !cs.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.ConnectionString = \"{}\"\n",
                        field_name, cs
                    ));
                }
            }
            code.push_str(&format!(
                "        Me.{}.Name = \"{}\"\n",
                field_name, control.name
            ));
            // Arbitrary properties for non-visual components (e.g. custom adapters)
            emit_arbitrary_props(&mut code, &field_name, control, &NON_VISUAL_HANDLED);
            continue;
        }

        // ── Visual control properties ─────────────────────────────────────

        // Text
        let text = control.get_text().unwrap_or(control.name.as_str());
        code.push_str(&format!("        Me.{}.Text = \"{}\"\n", field_name, text.replace('"', "\"\"")));

        // BackColor
        if let Some(bc) = control.get_back_color() {
            code.push_str(&format!(
                "        Me.{}.BackColor = System.Drawing.ColorTranslator.FromHtml(\"{}\")\n",
                field_name, bc
            ));
        } else if let Some(pv) = control.properties.get("BackColor") {
            // Named/system color stored as Expression
            if let Some(s) = property_value_to_vbnet(pv) {
                code.push_str(&format!("        Me.{}.BackColor = {}\n", field_name, s));
            }
        }

        // ForeColor
        if let Some(fc) = control.get_fore_color() {
            code.push_str(&format!(
                "        Me.{}.ForeColor = System.Drawing.ColorTranslator.FromHtml(\"{}\")\n",
                field_name, fc
            ));
        } else if let Some(pv) = control.properties.get("ForeColor") {
            if let Some(s) = property_value_to_vbnet(pv) {
                code.push_str(&format!("        Me.{}.ForeColor = {}\n", field_name, s));
            }
        }

        // Font
        if let Some(font_str) = control.get_font() {
            code.push_str(&format!(
                "        Me.{}.Font = {}\n",
                field_name,
                format_font(font_str)
            ));
        } else if let Some(pv) = control.properties.get("Font") {
            if let Some(s) = property_value_to_vbnet(pv) {
                code.push_str(&format!("        Me.{}.Font = {}\n", field_name, s));
            }
        }

        // Name and optional array Tag
        code.push_str(&format!(
            "        Me.{}.Name = \"{}\"\n",
            field_name, control.name
        ));
        if let Some(idx) = control.index {
            code.push_str(&format!(
                "        Me.{}.Tag = \"ArrayIndex={}\"\n",
                field_name, idx
            ));
        }

        // DataSource/DataMember for complex-binding controls (DataGridView, ListBox, ComboBox, BindingNavigator)
        if control.control_type.supports_complex_binding() {
            if let Some(ds) = control.properties.get_string("DataSource") {
                if !ds.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.DataSource = Me.{}\n",
                        field_name, ds
                    ));
                }
            }
            if let Some(dm) = control.properties.get_string("DataMember") {
                if !dm.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.DataMember = \"{}\"\n",
                        field_name, dm
                    ));
                }
            }
        }

        // DisplayMember/ValueMember for list controls
        if matches!(
            control.control_type,
            ControlType::ComboBox | ControlType::ListBox
        ) {
            if let Some(dpm) = control.properties.get_string("DisplayMember") {
                if !dpm.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.DisplayMember = \"{}\"\n",
                        field_name, dpm
                    ));
                }
            }
            if let Some(vm) = control.properties.get_string("ValueMember") {
                if !vm.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.ValueMember = \"{}\"\n",
                        field_name, vm
                    ));
                }
            }
        }

        // DataBindings.Add — emit all DataBindings.* prefix keys (not just the hardcoded four)
        if let Some(binding_source) = control
            .properties
            .get_string("DataBindings.Source")
            .map(|s| s.to_string())
        {
            if !binding_source.is_empty() {
                // Collect all DataBindings.<PropName> entries in stable order
                let mut db_entries: Vec<(&str, &str)> = control
                    .properties
                    .iter()
                    .filter_map(|(key, val)| {
                        if key.starts_with("DataBindings.") && key != "DataBindings.Source" {
                            let prop = &key["DataBindings.".len()..];
                            val.as_string().map(|col| (prop, col))
                        } else {
                            None
                        }
                    })
                    .collect();
                db_entries.sort_by_key(|(p, _)| *p);
                for (prop, col) in db_entries {
                    if !col.is_empty() {
                        code.push_str(&format!(
                            "        Me.{}.DataBindings.Add(\"{}\", Me.{}, \"{}\")\n",
                            field_name, prop, binding_source, col
                        ));
                    }
                }
            }
        }

        // BindingSource reference for BindingNavigator
        if matches!(control.control_type, ControlType::BindingNavigator) {
            if let Some(bs) = control.properties.get_string("BindingSource") {
                if !bs.is_empty() {
                    code.push_str(&format!(
                        "        Me.{}.BindingSource = Me.{}\n",
                        field_name, bs
                    ));
                }
            }
        }

        code.push_str(&format!(
            "        Me.{}.TabIndex = {}\n",
            field_name, control.tab_index
        ));

        // Arbitrary properties (everything not handled above)
        emit_arbitrary_props(&mut code, &field_name, control, &VISUAL_HANDLED);
    }

    // ── 3. Controls.Add — nested children first, then top-level ───────────
    // Children of containers first (Me.Panel.Controls.Add(Me.child))
    for control in &form.controls {
        if control.control_type.is_non_visual() {
            continue;
        }
        if let Some(parent_id) = control.parent_id {
            if let Some(parent) = form.controls.iter().find(|c| c.id == parent_id) {
                let parent_field = control_field_name(parent);
                let child_field = control_field_name(control);
                code.push_str(&format!(
                    "        Me.{}.Controls.Add(Me.{})\n",
                    parent_field, child_field
                ));
            }
        }
    }
    // Then top-level controls (directly on form)
    for control in &form.controls {
        if control.control_type.is_non_visual() {
            continue;
        }
        if control.parent_id.is_none() {
            let field_name = control_field_name(control);
            code.push_str(&format!("        Me.Controls.Add(Me.{})\n", field_name));
        }
    }

    // ── 4. Event wiring via AddHandler (designer-persisted bindings) ──────
    // These are emitted for events wired in InitializeComponent rather than via Handles clauses.
    for binding in &form.event_bindings {
        // Resolve the field name for the control (handles array members)
        let field_name = form
            .controls
            .iter()
            .find(|c| c.name.eq_ignore_ascii_case(&binding.control_name))
            .map(|c| control_field_name(c))
            .unwrap_or_else(|| binding.control_name.clone());
        code.push_str(&format!(
            "        AddHandler Me.{}.{}, AddressOf Me.{}\n",
            field_name,
            binding.event_type.as_str(),
            binding.handler_name
        ));
    }

    // ── 5. Form-level properties ───────────────────────────────────────────
    code.push_str(&format!(
        "        Me.ClientSize = New System.Drawing.Size({}, {})\n",
        form.width, form.height
    ));
    code.push_str(&format!(
        "        Me.Text = \"{}\"\n",
        form.text.replace('"', "\"\"")
    ));
    code.push_str(&format!("        Me.Name = \"{}\"\n", form.name));

    if let Some(bc) = &form.back_color {
        code.push_str(&format!(
            "        Me.BackColor = System.Drawing.ColorTranslator.FromHtml(\"{}\")\n",
            bc
        ));
    } else if let Some(pv) = form.properties.get("BackColor") {
        if let Some(s) = property_value_to_vbnet(pv) {
            code.push_str(&format!("        Me.BackColor = {}\n", s));
        }
    }

    if let Some(fc) = &form.fore_color {
        code.push_str(&format!(
            "        Me.ForeColor = System.Drawing.ColorTranslator.FromHtml(\"{}\")\n",
            fc
        ));
    } else if let Some(pv) = form.properties.get("ForeColor") {
        if let Some(s) = property_value_to_vbnet(pv) {
            code.push_str(&format!("        Me.ForeColor = {}\n", s));
        }
    }

    if let Some(font_str) = &form.font {
        code.push_str(&format!(
            "        Me.Font = {}\n",
            format_font(font_str)
        ));
    } else if let Some(pv) = form.properties.get("Font") {
        if let Some(s) = property_value_to_vbnet(pv) {
            code.push_str(&format!("        Me.Font = {}\n", s));
        }
    }

    // Arbitrary form-level properties (StartPosition, FormBorderStyle, etc.)
    let form_base_handled = ["BackColor", "ForeColor", "Font", "ClientSize", "Text", "Name"];
    let mut form_props: Vec<(&String, &PropertyValue)> = form
        .properties
        .iter()
        .filter(|(k, _)| !form_base_handled.contains(&k.as_str()))
        .collect();
    form_props.sort_by_key(|(k, _)| k.as_str());
    for (key, val) in form_props {
        if let Some(val_str) = property_value_to_vbnet(val) {
            code.push_str(&format!("        Me.{} = {}\n", key, val_str));
        }
    }

    code.push_str("        Me.ResumeLayout(False)\n");
    code.push_str("        Me.PerformLayout()\n");
    code.push_str("    End Sub\n");
    code.push_str("End Class\n");

    code
}

/// Properties that are explicitly handled in the visual control section.
/// These are excluded from the generic arbitrary-property pass to avoid duplicates.
static VISUAL_HANDLED: &[&str] = &[
    "Text", "BackColor", "ForeColor", "Font",
    "Name", "Tag", "TabIndex",
    // DataBindings keys (all DataBindings.* are handled dynamically)
    "DataBindings.Source",
    "DataBindings.Text", "DataBindings.Checked", "DataBindings.ImageLocation", "DataBindings.Value",
    // Data-source binding
    "DataSource", "DataMember", "DisplayMember", "ValueMember", "BindingSource",
    // Vybe-internal runtime keys with no VB.NET designer equivalent
    "List", "ListValues", "ListIndex", "ToolbarVisible",
    // CheckState and Value are stored internally; DropDownStyle stored as int but VB.NET expects enum
    "CheckState", "DropDownStyle",
];

/// Properties handled in the non-visual component section.
static NON_VISUAL_HANDLED: &[&str] = &[
    "DataSource", "DataMember", "Filter", "Sort", "TableName", "DataSetName",
    "SelectCommand", "ConnectionString", "Name",
];

/// Emit all properties NOT in the `handled` list as generic VB.NET assignments.
/// Skips `DataBindings.*` prefix keys (emitted separately as DataBindings.Add calls).
/// Skips StringArray values (need special container-specific syntax).
fn emit_arbitrary_props(
    code: &mut String,
    field_name: &str,
    control: &Control,
    handled: &[&str],
) {
    // Sort for deterministic output
    let mut props: Vec<(&String, &PropertyValue)> = control
        .properties
        .iter()
        .filter(|(key, _)| {
            // Skip handled properties
            if handled.contains(&key.as_str()) {
                // Special case: Tag should still be emitted if it's NOT an array-index tag
                if *key == "Tag" {
                    return control.index.is_none(); // only emit if not an array member
                }
                return false;
            }
            // Skip DataBindings.* — handled by DataBindings.Add
            if key.starts_with("DataBindings.") {
                return false;
            }
            true
        })
        .collect();
    props.sort_by_key(|(k, _)| k.as_str());

    for (key, val) in props {
        if let Some(val_str) = property_value_to_vbnet(val) {
            code.push_str(&format!(
                "        Me.{}.{} = {}\n",
                field_name, key, val_str
            ));
        }
    }
}

/// Generates a minimal user code stub for a new VB.NET form.
pub fn generate_user_code_stub(form_name: &str) -> String {
    format!("Partial Class {}\n\nEnd Class\n", form_name)
}
