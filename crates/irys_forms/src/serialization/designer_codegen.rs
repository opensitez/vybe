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

        code.push_str(&format!(
            "        Me.{}.Location = New System.Drawing.Point({}, {})\n",
            field_name, control.bounds.x, control.bounds.y
        ));
        code.push_str(&format!(
            "        Me.{}.Size = New System.Drawing.Size({}, {})\n",
            field_name, control.bounds.width, control.bounds.height
        ));

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
