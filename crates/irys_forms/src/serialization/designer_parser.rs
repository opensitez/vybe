use crate::control::{Control, ControlType};
use crate::form::Form;
use irys_parser::{ClassDecl, Expression, Statement};

/// Extracts the last component of a potentially dotted identifier.
/// e.g. "System.Windows.Forms.Button" -> "Button", "Button" -> "Button"
fn last_component(name: &str) -> &str {
    name.rsplit('.').next().unwrap_or(name)
}

/// Maps a VB.NET type name to a ControlType (case-insensitive).
/// Handles both simple ("Button") and fully-qualified ("System.Windows.Forms.Button") names.
pub fn vbnet_type_to_control_type(name: &str) -> Option<ControlType> {
    match last_component(name).to_lowercase().as_str() {
        "button" => Some(ControlType::Button),
        "label" => Some(ControlType::Label),
        "textbox" => Some(ControlType::TextBox),
        "checkbox" => Some(ControlType::CheckBox),
        "radiobutton" => Some(ControlType::RadioButton),
        "combobox" => Some(ControlType::ComboBox),
        "listbox" => Some(ControlType::ListBox),
        "groupbox" => Some(ControlType::Frame),
        "picturebox" => Some(ControlType::PictureBox),
        "richtextbox" => Some(ControlType::RichTextBox),
        "webbrowser" => Some(ControlType::WebBrowser),
        "treeview" => Some(ControlType::TreeView),
        "datagridview" => Some(ControlType::DataGridView),
        "panel" => Some(ControlType::Panel),
        "listview" => Some(ControlType::ListView),
        "bindingnavigator" => Some(ControlType::BindingNavigator),
        "bindingsource" => Some(ControlType::BindingSourceComponent),
        "dataset" => Some(ControlType::DataSetComponent),
        "datatable" => Some(ControlType::DataTableComponent),
        "sqldataadapter" | "dataadapter" | "oledbdataadapter" => Some(ControlType::DataAdapterComponent),
        _ => None,
    }
}

/// Accumulator for building a Control from designer AST statements.
struct ControlBuilder {
    name: String,
    control_type: ControlType,
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    text: Option<String>,
    back_color: Option<String>,
    fore_color: Option<String>,
    font: Option<String>,
    tab_index: i32,
    tag: Option<String>,
    explicit_name: Option<String>,
    /// Additional string properties (DataSource, DataMember, DisplayMember, etc.)
    extra_props: std::collections::HashMap<String, String>,
}

impl ControlBuilder {
    fn new(name: String, control_type: ControlType) -> Self {
        let (w, h) = control_type.default_size();
        Self {
            name,
            control_type,
            x: 0,
            y: 0,
            width: w,
            height: h,
            text: None,
            back_color: None,
            fore_color: None,
            font: None,
            tab_index: 0,
            tag: None,
            explicit_name: None,
            extra_props: std::collections::HashMap::new(),
        }
    }

    fn build(self) -> Control {
        // Use explicit Name property if set (recovers base name for array members)
        let control_name = self.explicit_name.unwrap_or(self.name);
        let mut ctrl = Control::new(self.control_type, control_name, self.x, self.y);
        ctrl.bounds.width = self.width;
        ctrl.bounds.height = self.height;
        ctrl.tab_index = self.tab_index;
        // Check Tag for "ArrayIndex=N" pattern
        if let Some(ref tag) = self.tag {
            if let Some(idx_str) = tag.strip_prefix("ArrayIndex=") {
                if let Ok(idx) = idx_str.parse::<i32>() {
                    ctrl.index = Some(idx);
                }
            }
        }
        if let Some(text) = self.text {
            // In .NET WinForms, ALL controls use the Text property (Caption is VB6-only)
            ctrl.set_text(text);
        }
        if let Some(bc) = self.back_color {
            ctrl.set_back_color(bc);
        }
        if let Some(fc) = self.fore_color {
            ctrl.set_fore_color(fc);
        }
        if let Some(font) = self.font {
            ctrl.set_font(font);
        }
        // Apply extra data-binding properties
        for (key, val) in &self.extra_props {
            ctrl.properties.set(key.as_str(), val.clone());
        }
        ctrl
    }
}

/// Extracts a Form object from a merged ClassDecl by analyzing InitializeComponent.
///
/// Looks for a Sub named "InitializeComponent" in the class methods, then walks
/// the AST statements to reconstruct controls and form properties.
pub fn extract_form_from_designer(class_decl: &ClassDecl) -> Option<Form> {
    // Find InitializeComponent method
    let init_method = class_decl.methods.iter().find_map(|m| {
        match m {
            irys_parser::MethodDecl::Sub(s) if s.name.as_str().eq_ignore_ascii_case("InitializeComponent") => {
                Some(&s.body)
            }
            _ => None,
        }
    })?;

    let form_name = class_decl.name.as_str().to_string();
    let mut form = Form::new(&form_name);

    // Use IndexMap to preserve insertion order for deterministic control ordering
    let mut builders: std::collections::HashMap<String, ControlBuilder> = std::collections::HashMap::new();
    let mut insertion_order: Vec<String> = Vec::new();

    for stmt in init_method {
        match stmt {
            // Me.X = value -> form property or control registration
            Statement::MemberAssignment { object, member, value } if is_me(object) => {
                let member_name = member.as_str();

                // Check form-level properties first (before control registration)
                if member_name.eq_ignore_ascii_case("ClientSize") {
                    if let Some((w, h)) = extract_size(value) {
                        form.width = w;
                        form.height = h;
                    }
                } else if member_name.eq_ignore_ascii_case("Text") {
                    if let Expression::StringLiteral(s) = value {
                        form.text = s.clone();
                    }
                } else if member_name.eq_ignore_ascii_case("Name") {
                    if let Expression::StringLiteral(s) = value {
                        form.name = s.clone();
                    }
                } else if member_name.eq_ignore_ascii_case("BackColor") {
                    if let Some(color) = extract_color(value) {
                        form.back_color = Some(color);
                    }
                } else if member_name.eq_ignore_ascii_case("ForeColor") {
                    if let Some(color) = extract_color(value) {
                        form.fore_color = Some(color);
                    }
                } else if member_name.eq_ignore_ascii_case("Font") {
                    if let Some(font) = extract_font(value) {
                        form.font = Some(font);
                    }
                } else if let Expression::New(type_id, _) = value {
                    // Me.X = New System.Windows.Forms.Button() -> register control
                    if let Some(ct) = vbnet_type_to_control_type(type_id.as_str()) {
                        if !builders.contains_key(member_name) {
                            insertion_order.push(member_name.to_string());
                        }
                        builders.insert(
                            member_name.to_string(),
                            ControlBuilder::new(member_name.to_string(), ct),
                        );
                    }
                }
            }
            // Me.X.Prop = value -> set control property
            Statement::MemberAssignment { object, member, value } => {
                if let Some((ctrl_name, is_me_prefix)) = extract_me_member_target(object) {
                    if !is_me_prefix {
                        continue;
                    }
                    let prop_name = member.as_str();
                    if let Some(builder) = builders.get_mut(&ctrl_name) {
                        match prop_name.to_lowercase().as_str() {
                            "location" => {
                                if let Some((x, y)) = extract_point(value) {
                                    builder.x = x;
                                    builder.y = y;
                                }
                            }
                            "size" => {
                                if let Some((w, h)) = extract_size(value) {
                                    builder.width = w;
                                    builder.height = h;
                                }
                            }
                            "text" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.text = Some(s.clone());
                                }
                            }
                            "backcolor" => {
                                if let Some(color) = extract_color(value) {
                                    builder.back_color = Some(color);
                                }
                            }
                            "forecolor" => {
                                if let Some(color) = extract_color(value) {
                                    builder.fore_color = Some(color);
                                }
                            }
                            "font" => {
                                if let Some(font) = extract_font(value) {
                                    builder.font = Some(font);
                                }
                            }
                            "name" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.explicit_name = Some(s.clone());
                                }
                            }
                            "tag" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.tag = Some(s.clone());
                                }
                            }
                            "tabindex" => {
                                if let Expression::IntegerLiteral(n) = value {
                                    builder.tab_index = *n;
                                }
                            }
                            // Data binding properties: Me.ctrl.DataSource = Me.bs1
                            "datasource" => {
                                match value {
                                    Expression::MemberAccess(inner, member) if is_me(inner) => {
                                        builder.extra_props.insert("DataSource".to_string(), member.as_str().to_string());
                                    }
                                    Expression::StringLiteral(s) => {
                                        builder.extra_props.insert("DataSource".to_string(), s.clone());
                                    }
                                    _ => {}
                                }
                            }
                            "datamember" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.extra_props.insert("DataMember".to_string(), s.clone());
                                }
                            }
                            "displaymember" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.extra_props.insert("DisplayMember".to_string(), s.clone());
                                }
                            }
                            "valuemember" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.extra_props.insert("ValueMember".to_string(), s.clone());
                                }
                            }
                            "bindingsource" => {
                                match value {
                                    Expression::MemberAccess(inner, member) if is_me(inner) => {
                                        builder.extra_props.insert("BindingSource".to_string(), member.as_str().to_string());
                                    }
                                    Expression::StringLiteral(s) => {
                                        builder.extra_props.insert("BindingSource".to_string(), s.clone());
                                    }
                                    _ => {}
                                }
                            }
                            "filter" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.extra_props.insert("Filter".to_string(), s.clone());
                                }
                            }
                            "sort" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.extra_props.insert("Sort".to_string(), s.clone());
                                }
                            }
                            "datasetname" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.extra_props.insert("DataSetName".to_string(), s.clone());
                                }
                            }
                            "tablename" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.extra_props.insert("TableName".to_string(), s.clone());
                                }
                            }
                            "selectcommand" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.extra_props.insert("SelectCommand".to_string(), s.clone());
                                }
                            }
                            "connectionstring" => {
                                if let Expression::StringLiteral(s) = value {
                                    builder.extra_props.insert("ConnectionString".to_string(), s.clone());
                                }
                            }
                            _ => {}
                        }
                    }
                }
            }
            // Handle DataBindings.Add calls:
            // Me.txtName.DataBindings.Add("Text", Me.bs1, "Name")
            // Parsed as ExpressionStatement(MethodCall(MemberAccess(MemberAccess(Me, ctrl), "DataBindings"), "Add", args))
            Statement::ExpressionStatement(Expression::MethodCall(obj, method, args))
                if method.as_str().eq_ignore_ascii_case("Add") =>
            {
                // Check if the object is Me.X.DataBindings
                if let Expression::MemberAccess(inner, db_member) = obj.as_ref() {
                    if db_member.as_str().eq_ignore_ascii_case("DataBindings") {
                        if let Some((ctrl_name, true)) = extract_me_member_target(inner) {
                            // args: ("PropertyName", Me.bindingSource, "ColumnName")
                            if args.len() >= 3 {
                                let prop_name = match &args[0] {
                                    Expression::StringLiteral(s) => Some(s.clone()),
                                    _ => None,
                                };
                                let bs_name = match &args[1] {
                                    Expression::MemberAccess(inner2, member) if is_me(inner2) => {
                                        Some(member.as_str().to_string())
                                    }
                                    Expression::StringLiteral(s) => Some(s.clone()),
                                    _ => None,
                                };
                                let col_name = match &args[2] {
                                    Expression::StringLiteral(s) => Some(s.clone()),
                                    _ => None,
                                };
                                if let (Some(prop), Some(bs), Some(col)) = (prop_name, bs_name, col_name) {
                                    if let Some(builder) = builders.get_mut(&ctrl_name) {
                                        builder.extra_props.insert("DataBindings.Source".to_string(), bs);
                                        builder.extra_props.insert(format!("DataBindings.{}", prop), col);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            // Me.Controls.Add(...), Me.SuspendLayout(), Me.ResumeLayout() -> skip
            Statement::ExpressionStatement(_) | Statement::Call { .. } => {}
            _ => {}
        }
    }

    // Build controls in insertion order
    for name in &insertion_order {
        if let Some(builder) = builders.remove(name) {
            form.controls.push(builder.build());
        }
    }

    Some(form)
}

/// Check if an expression is just `Me`
fn is_me(expr: &Expression) -> bool {
    matches!(expr, Expression::Me)
}

/// Given an expression like MemberAccess(Me, "Button1"), return ("Button1", true).
fn extract_me_member_target(expr: &Expression) -> Option<(String, bool)> {
    match expr {
        Expression::MemberAccess(inner, member) => {
            if is_me(inner) {
                Some((member.as_str().to_string(), true))
            } else {
                None
            }
        }
        _ => None,
    }
}

/// Extract (x, y) from New Point(x, y) or New System.Drawing.Point(x, y)
fn extract_point(expr: &Expression) -> Option<(i32, i32)> {
    if let Expression::New(id, args) = expr {
        if last_component(id.as_str()).eq_ignore_ascii_case("Point") && args.len() == 2 {
            let x = expr_to_i32(&args[0])?;
            let y = expr_to_i32(&args[1])?;
            return Some((x, y));
        }
    }
    None
}

/// Extract (w, h) from New Size(w, h) or New System.Drawing.Size(w, h)
fn extract_size(expr: &Expression) -> Option<(i32, i32)> {
    if let Expression::New(id, args) = expr {
        if last_component(id.as_str()).eq_ignore_ascii_case("Size") && args.len() == 2 {
            let w = expr_to_i32(&args[0])?;
            let h = expr_to_i32(&args[1])?;
            return Some((w, h));
        }
    }
    None
}

/// Extract color hex string (#RRGGBB) from common VB.NET designer color expressions.
fn extract_color(expr: &Expression) -> Option<String> {
    match expr {
        // ColorTranslator.FromHtml("#RRGGBB")
        Expression::Call(ident, args) => {
            let name = ident.as_str();
            if name.eq_ignore_ascii_case("FromHtml") && args.len() == 1 {
                if let Expression::StringLiteral(s) = &args[0] {
                    return Some(s.clone());
                }
            }
            if name.eq_ignore_ascii_case("FromArgb") {
                if args.len() == 1 {
                    if let Some(val) = expr_to_i32(&args[0]) {
                        let v = val as u32;
                        let r = ((v >> 16) & 0xFF) as u8;
                        let g = ((v >> 8) & 0xFF) as u8;
                        let b = (v & 0xFF) as u8;
                        return Some(format!("#{:02X}{:02X}{:02X}", r, g, b));
                    }
                } else if args.len() == 3 || args.len() == 4 {
                    let start = if args.len() == 4 { 1 } else { 0 }; // skip alpha if present
                    let r = expr_to_i32(&args[start])? as u8;
                    let g = expr_to_i32(&args[start + 1])? as u8;
                    let b = expr_to_i32(&args[start + 2])? as u8;
                    return Some(format!("#{:02X}{:02X}{:02X}", r, g, b));
                }
            }
            None
        }
        Expression::MethodCall(_, ident, args) => {
            let name = ident.as_str();
            if name.eq_ignore_ascii_case("FromHtml") && args.len() == 1 {
                if let Expression::StringLiteral(s) = &args[0] {
                    return Some(s.clone());
                }
            }
            if name.eq_ignore_ascii_case("FromArgb") {
                if args.len() == 1 {
                    if let Some(val) = expr_to_i32(&args[0]) {
                        let v = val as u32;
                        let r = ((v >> 16) & 0xFF) as u8;
                        let g = ((v >> 8) & 0xFF) as u8;
                        let b = (v & 0xFF) as u8;
                        return Some(format!("#{:02X}{:02X}{:02X}", r, g, b));
                    }
                } else if args.len() == 3 || args.len() == 4 {
                    let start = if args.len() == 4 { 1 } else { 0 };
                    let r = expr_to_i32(&args[start])? as u8;
                    let g = expr_to_i32(&args[start + 1])? as u8;
                    let b = expr_to_i32(&args[start + 2])? as u8;
                    return Some(format!("#{:02X}{:02X}{:02X}", r, g, b));
                }
            }
            None
        }
        _ => None,
    }
}

/// Extract font string "Family, sizepx" from New Font(...) expressions.
fn extract_font(expr: &Expression) -> Option<String> {
    if let Expression::New(id, args) = expr {
        if last_component(id.as_str()).eq_ignore_ascii_case("Font") {
            if args.len() >= 2 {
                if let Expression::StringLiteral(fam) = &args[0] {
                    let size = match expr_to_i32(&args[1]) {
                        Some(i) => i as f32,
                        None => match &args[1] {
                            Expression::DoubleLiteral(d) => *d as f32,
                            _ => 12.0,
                        },
                    };
                    return Some(format!("{}, {}px", fam, size));
                }
            }
        }
    }
    None
}

fn expr_to_i32(expr: &Expression) -> Option<i32> {
    match expr {
        Expression::IntegerLiteral(n) => Some(*n),
        Expression::DoubleLiteral(d) => Some(*d as i32),
        _ => None,
    }
}
