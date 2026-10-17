//! C# WinForms designer CODE GENERATION.
//!
//! ⚠ Everything below `generate_designer_code` / `generate_user_code_stub` is
//! now DEAD: the `load_designer` / `save_designer` pair that used those helpers
//! was registered on `FormModuleLanguage` and never invoked, and has been
//! removed. The helpers are left for a separate dead-code pass rather than
//! folded into an unrelated change.

use vybe_ast::*;
use vybe_platform_dotnet::winforms::form::Form;

/// Emit C# designer code from the shared form model.
pub fn generate_designer_code(form: &Form) -> String {
    let mut out = String::new();
    out.push_str(&format!(
        "public partial class {} : System.Windows.Forms.Form\n{{\n",
        form.name
    ));

    for control in &form.controls {
        let ty = control.control_type.dotnet_class_name();
        out.push_str(&format!("    private {} {};\n", ty, control.name));
    }

    out.push_str("\n    private void InitializeComponent()\n    {\n");
    out.push_str("        this.SuspendLayout();\n");
    for control in &form.controls {
        let ty = control.control_type.dotnet_class_name();
        out.push_str(&format!(
            "        this.{0} = new {1}();\n",
            control.name, ty
        ));
        out.push_str(&format!(
            "        this.{0}.Location = new System.Drawing.Point({1}, {2});\n",
            control.name, control.bounds.x, control.bounds.y
        ));
        out.push_str(&format!(
            "        this.{0}.Size = new System.Drawing.Size({1}, {2});\n",
            control.name, control.bounds.width, control.bounds.height
        ));
        out.push_str(&format!("        this.{0}.Name = \"{0}\";\n", control.name));
        if let Some(text) = control.get_text() {
            out.push_str(&format!(
                "        this.{0}.Text = \"{1}\";\n",
                control.name,
                text.replace('"', "\\\"")
            ));
        }
        out.push_str(&format!(
            "        this.Controls.Add(this.{});\n",
            control.name
        ));
    }
    out.push_str(&format!(
        "        this.ClientSize = new System.Drawing.Size({}, {});\n",
        form.width, form.height
    ));
    out.push_str(&format!(
        "        this.Text = \"{}\";\n",
        form.text.replace('"', "\\\"")
    ));
    out.push_str("        this.ResumeLayout(false);\n");
    out.push_str("    }\n}\n");
    out
}

/// Generate a minimal user-code stub for a C# form partial class.
pub fn generate_user_code_stub(form_name: &str) -> String {
    format!(
        "public partial class {0} : System.Windows.Forms.Form\n{{\n    public {0}()\n    {{\n        InitializeComponent();\n    }}\n}}\n",
        form_name
    )
}















