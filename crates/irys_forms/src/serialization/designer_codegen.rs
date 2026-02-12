use crate::control::{Control, ControlType};
use crate::form::Form;

/// Returns the field name for a control in VB.NET designer code.
/// Array members use `Name_Index` (e.g., `Command1_0`), non-array use `Name`.
fn control_field_name(control: &Control) -> String {
    if let Some(idx) = control.index {
        format!("{}_{}", control.name, idx)
    } else {
        control.name.clone()
    }
}

/// Maps a ControlType to its fully-qualified VB.NET class name
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
        ControlType::BindingSourceComponent => "System.Windows.Forms.BindingSource",
        ControlType::DataSetComponent => "System.Data.DataSet",
        ControlType::DataTableComponent => "System.Data.DataTable",
        ControlType::DataAdapterComponent => "System.Data.SqlClient.SqlDataAdapter",
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
        code.push_str(&format!("    Friend WithEvents {} As {}\n", field_name, vb_type));
    }
    code.push('\n');

    code.push_str("    Private Sub InitializeComponent()\n");

    // Control instantiation
    for control in &form.controls {
        let vb_type = control_type_to_vbnet(&control.control_type);
        let field_name = control_field_name(control);
        code.push_str(&format!("        Me.{} = New {}()\n", field_name, vb_type));
    }

    code.push_str("        Me.SuspendLayout()\n");

    // Control property assignment
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
            if let Some(ds) = control.properties.get_string("DataSource") {
                if !ds.is_empty() {
                    code.push_str(&format!("        Me.{}.DataSource = Me.{}\n", field_name, ds));
                }
            }
            if let Some(dm) = control.properties.get_string("DataMember") {
                if !dm.is_empty() {
                    code.push_str(&format!("        Me.{}.DataMember = \"{}\"\n", field_name, dm));
                }
            }
            if let Some(filter) = control.properties.get_string("Filter") {
                if !filter.is_empty() {
                    code.push_str(&format!("        Me.{}.Filter = \"{}\"\n", field_name, filter));
                }
            }
            if let Some(sort) = control.properties.get_string("Sort") {
                if !sort.is_empty() {
                    code.push_str(&format!("        Me.{}.Sort = \"{}\"\n", field_name, sort));
                }
            }
            if let Some(tn) = control.properties.get_string("TableName") {
                if !tn.is_empty() {
                    code.push_str(&format!("        Me.{}.TableName = \"{}\"\n", field_name, tn));
                }
            }
            if let Some(dsn) = control.properties.get_string("DataSetName") {
                if !dsn.is_empty() {
                    code.push_str(&format!("        Me.{}.DataSetName = \"{}\"\n", field_name, dsn));
                }
            }
            if let Some(sc) = control.properties.get_string("SelectCommand") {
                if !sc.is_empty() {
                    code.push_str(&format!("        Me.{}.SelectCommand = \"{}\"\n", field_name, sc));
                }
            }
            if let Some(cs) = control.properties.get_string("ConnectionString") {
                if !cs.is_empty() {
                    code.push_str(&format!("        Me.{}.ConnectionString = \"{}\"\n", field_name, cs));
                }
            }
            code.push_str(&format!("        Me.{}.Name = \"{}\"\n", field_name, control.name));
            continue;
        }

        // Text/Caption property
        let text = control
            .get_text()
            .or_else(|| control.get_caption())
            .unwrap_or(control.name.as_str());
        code.push_str(&format!("        Me.{}.Text = \"{}\"\n", field_name, text));

        // Colors
        if let Some(bc) = control.get_back_color() {
            code.push_str(&format!(
                "        Me.{}.BackColor = System.Drawing.ColorTranslator.FromHtml(\"{}\")\n",
                field_name, bc
            ));
        }
        if let Some(fc) = control.get_fore_color() {
            code.push_str(&format!(
                "        Me.{}.ForeColor = System.Drawing.ColorTranslator.FromHtml(\"{}\")\n",
                field_name, fc
            ));
        }

        // Font (expects format: Family, size[px/pt])
        if let Some(font_str) = control.get_font() {
            let mut parts = font_str.split(',').map(|s| s.trim());
            let family = parts.next().unwrap_or("Segoe UI");
            let size_raw = parts.next().unwrap_or("12");
            let size_clean = size_raw.trim_end_matches("px").trim_end_matches("pt");
            let size_val: f32 = size_clean.parse().unwrap_or(12.0);
            code.push_str(&format!(
                "        Me.{}.Font = New System.Drawing.Font(\"{}\", {}F)\n",
                field_name,
                family.replace("\"", "\"\""),
                size_val
            ));
        }

        // For array members, set Name to base name and Tag with array index
        code.push_str(&format!("        Me.{}.Name = \"{}\"\n", field_name, control.name));
        if let Some(idx) = control.index {
            code.push_str(&format!("        Me.{}.Tag = \"ArrayIndex={}\"\n", field_name, idx));
        }

        // DataSource binding for data-bound visual controls
        if control.control_type.supports_complex_binding() {
            if let Some(ds) = control.properties.get_string("DataSource") {
                if !ds.is_empty() {
                    code.push_str(&format!("        Me.{}.DataSource = Me.{}\n", field_name, ds));
                }
            }
            if let Some(dm) = control.properties.get_string("DataMember") {
                if !dm.is_empty() {
                    code.push_str(&format!("        Me.{}.DataMember = \"{}\"\n", field_name, dm));
                }
            }
        }

        // DisplayMember/ValueMember for list controls (ComboBox, ListBox)
        if matches!(control.control_type, ControlType::ComboBox | ControlType::ListBox) {
            if let Some(dpm) = control.properties.get_string("DisplayMember") {
                if !dpm.is_empty() {
                    code.push_str(&format!("        Me.{}.DisplayMember = \"{}\"\n", field_name, dpm));
                }
            }
            if let Some(vm) = control.properties.get_string("ValueMember") {
                if !vm.is_empty() {
                    code.push_str(&format!("        Me.{}.ValueMember = \"{}\"\n", field_name, vm));
                }
            }
        }

        // Simple data bindings (DataBindings.Add) for all visual controls
        if let Some(binding_source) = control.properties.get_string("DataBindings.Source") {
            if !binding_source.is_empty() {
                // Determine which property is being bound
                let bindable_props = ["Text", "Checked", "ImageLocation", "Value"];
                for prop in &bindable_props {
                    let key = format!("DataBindings.{}", prop);
                    if let Some(col) = control.properties.get_string(&key) {
                        if !col.is_empty() {
                            code.push_str(&format!(
                                "        Me.{}.DataBindings.Add(\"{}\", Me.{}, \"{}\")\n",
                                field_name, prop, binding_source, col
                            ));
                        }
                    }
                }
            }
        }

        // BindingSource reference for BindingNavigator
        if matches!(control.control_type, ControlType::BindingNavigator) {
            if let Some(bs) = control.properties.get_string("BindingSource") {
                if !bs.is_empty() {
                    code.push_str(&format!("        Me.{}.BindingSource = Me.{}\n", field_name, bs));
                }
            }
        }

        code.push_str(&format!(
            "        Me.{}.TabIndex = {}\n",
            field_name, control.tab_index
        ));
        code.push_str(&format!("        Me.Controls.Add(Me.{})\n", field_name));
    }

    // Form properties
    code.push_str(&format!(
        "        Me.ClientSize = New System.Drawing.Size({}, {})\n",
        form.width, form.height
    ));
    code.push_str(&format!("        Me.Text = \"{}\"\n", form.caption));
    code.push_str(&format!("        Me.Name = \"{}\"\n", form.name));
    if let Some(bc) = &form.back_color {
        code.push_str(&format!(
            "        Me.BackColor = System.Drawing.ColorTranslator.FromHtml(\"{}\")\n",
            bc
        ));
    }
    if let Some(fc) = &form.fore_color {
        code.push_str(&format!(
            "        Me.ForeColor = System.Drawing.ColorTranslator.FromHtml(\"{}\")\n",
            fc
        ));
    }
    if let Some(font_str) = &form.font {
        let mut parts = font_str.split(',').map(|s| s.trim());
        let family = parts.next().unwrap_or("Segoe UI");
        let size_raw = parts.next().unwrap_or("12");
        let size_clean = size_raw.trim_end_matches("px").trim_end_matches("pt");
        let size_val: f32 = size_clean.parse().unwrap_or(12.0);
        code.push_str(&format!(
            "        Me.Font = New System.Drawing.Font(\"{}\", {}F)\n",
            family.replace("\"", "\"\""),
            size_val
        ));
    }
    code.push_str("        Me.ResumeLayout(False)\n");

    code.push_str("    End Sub\n");
    code.push_str("End Class\n");

    code
}

/// Generates a minimal user code stub for a new VB.NET form.
pub fn generate_user_code_stub(form_name: &str) -> String {
    format!("Partial Class {}\n\nEnd Class\n", form_name)
}
