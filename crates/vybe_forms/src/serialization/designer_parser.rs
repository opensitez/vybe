use crate::control::{Control, ControlType};
use crate::events::{EventBinding, EventType};
use crate::form::Form;
use vybe_parser::{ClassDecl, Expression, Statement};

/// Extracts the last component of a potentially dotted identifier.
/// e.g. "System.Windows.Forms.Button" -> "Button", "Button" -> "Button"
fn last_component(name: &str) -> &str {
    name.rsplit('.').next().unwrap_or(name)
}

/// Maps a VB.NET type name to a ControlType (case-insensitive).
/// Handles both simple ("Button") and fully-qualified ("System.Windows.Forms.Button") names.
/// Unknown types become Custom(fully_qualified_name) for round-trip fidelity.
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
        "tabcontrol" => Some(ControlType::TabControl),
        "tabpage" => Some(ControlType::TabPage),
        "progressbar" => Some(ControlType::ProgressBar),
        "numericupdown" => Some(ControlType::NumericUpDown),
        "menustrip" => Some(ControlType::MenuStrip),
        "toolstripmenuitem" => Some(ControlType::ToolStripMenuItem),
        "contextmenustrip" => Some(ControlType::ContextMenuStrip),
        "statusstrip" => Some(ControlType::StatusStrip),
        "toolstripstatuslabel" => Some(ControlType::ToolStripStatusLabel),
        "datetimepicker" => Some(ControlType::DateTimePicker),
        "linklabel" => Some(ControlType::LinkLabel),
        "toolstrip" => Some(ControlType::ToolStrip),
        "trackbar" => Some(ControlType::TrackBar),
        "maskedtextbox" => Some(ControlType::MaskedTextBox),
        "splitcontainer" => Some(ControlType::SplitContainer),
        "flowlayoutpanel" => Some(ControlType::FlowLayoutPanel),
        "tablelayoutpanel" => Some(ControlType::TableLayoutPanel),
        "monthcalendar" => Some(ControlType::MonthCalendar),
        "hscrollbar" => Some(ControlType::HScrollBar),
        "vscrollbar" => Some(ControlType::VScrollBar),
        "tooltip" => Some(ControlType::ToolTip),
        "bindingsource" => Some(ControlType::BindingSourceComponent),
        "dataset" => Some(ControlType::DataSetComponent),
        "datatable" => Some(ControlType::DataTableComponent),
        "sqldataadapter" | "dataadapter" | "oledbdataadapter" => Some(ControlType::DataAdapterComponent),
        // Non-visual infrastructure components
        "timer" => Some(ControlType::Timer),
        "imagelist" => Some(ControlType::ImageList),
        "errorprovider" => Some(ControlType::ErrorProvider),
        // Dialog components
        "openfiledialog" => Some(ControlType::OpenFileDialog),
        "savefiledialog" => Some(ControlType::SaveFileDialog),
        "folderbrowserdialog" => Some(ControlType::FolderBrowserDialog),
        "fontdialog" => Some(ControlType::FontDialog),
        "colordialog" => Some(ControlType::ColorDialog),
        "printdialog" => Some(ControlType::PrintDialog),
        "printdocument" => Some(ControlType::PrintDocument),
        // Notification / system tray
        "notifyicon" => Some(ControlType::NotifyIcon),
        // Additional visual controls
        "checkedlistbox" => Some(ControlType::CheckedListBox),
        "domainupdown" => Some(ControlType::DomainUpDown),
        "propertygrid" => Some(ControlType::PropertyGrid),
        "splitter" => Some(ControlType::Splitter),
        "datagrid" => Some(ControlType::DataGrid),
        "usercontrol" => Some(ControlType::UserControl),
        // ToolStrip sub-components
        "toolstripseparator" => Some(ControlType::ToolStripSeparator),
        "toolstripbutton" => Some(ControlType::ToolStripButton),
        "toolstriplabel" => Some(ControlType::ToolStripLabel),
        "toolstripcombobox" => Some(ControlType::ToolStripComboBox),
        "toolstripdropdownbutton" => Some(ControlType::ToolStripDropDownButton),
        "toolstripsplitbutton" => Some(ControlType::ToolStripSplitButton),
        "toolstriptextbox" => Some(ControlType::ToolStripTextBox),
        "toolstripprogressbar" => Some(ControlType::ToolStripProgressBar),
        // Additional dialogs (non-visual)
        "printpreviewdialog" => Some(ControlType::PrintPreviewDialog),
        "pagesetupdialog" => Some(ControlType::PageSetupDialog),
        "printpreviewcontrol" => Some(ControlType::PrintPreviewControl),
        // Non-visual infrastructure
        "helpprovider" => Some(ControlType::HelpProvider),
        "backgroundworker" => Some(ControlType::BackgroundWorker),
        "sqlconnection" => Some(ControlType::SqlConnection),
        "oledbconnection" => Some(ControlType::OleDbConnection),
        "dataview" => Some(ControlType::DataView),
        _ => Some(ControlType::Custom(name.to_string())),
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
    /// Additional properties (DataSource, DataMember, etc.) and arbitrary designer properties.
    /// Keys use original casing from the designer file for round-trip fidelity.
    extra_props: std::collections::HashMap<String, crate::properties::PropertyValue>,
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
        // Use explicit Name property if set (recovers base name for array members like "btn1_0" → "btn1")
        let control_name = self.explicit_name.unwrap_or(self.name);
        let mut ctrl = Control::new(self.control_type, control_name, self.x, self.y);
        ctrl.bounds.width = self.width;
        ctrl.bounds.height = self.height;
        ctrl.tab_index = self.tab_index;
        // Check Tag for "ArrayIndex=N" pattern (VB6-style control arrays)
        if let Some(ref tag) = self.tag {
            if let Some(idx_str) = tag.strip_prefix("ArrayIndex=") {
                if let Ok(idx) = idx_str.parse::<i32>() {
                    ctrl.index = Some(idx);
                }
            }
            // Always persist the raw Tag value to properties for round-trip fidelity
            ctrl.properties.set("Tag", tag.clone());
        }
        if let Some(text) = self.text {
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
        for (key, val) in self.extra_props {
            ctrl.properties.set_raw(key, val);
        }
        ctrl
    }
}

/// Convert an AST expression to a PropertyValue.
/// Literals become typed values; everything else becomes an Expression code string
/// so round-trip fidelity is preserved for complex expressions.
fn expr_to_property_value(expr: &Expression) -> crate::properties::PropertyValue {
    use crate::properties::PropertyValue;
    match expr {
        Expression::StringLiteral(s) => PropertyValue::String(s.clone()),
        Expression::IntegerLiteral(i) => PropertyValue::Integer(*i),
        Expression::BooleanLiteral(b) => PropertyValue::Boolean(*b),
        Expression::DoubleLiteral(d) => PropertyValue::Double(*d),
        _ => PropertyValue::Expression(expr_to_code(expr)),
    }
}

/// Reconstruct a VB.NET code string from an AST expression.
/// Used for preserving complex property values (images, enums, etc.) verbatim.
fn expr_to_code(expr: &Expression) -> String {
    match expr {
        Expression::Variable(id) => id.as_str().to_string(),
        Expression::MemberAccess(obj, member) => {
            format!("{}.{}", expr_to_code(obj), member.as_str())
        }
        Expression::Call(func, args) => {
            let arg_strs: Vec<String> = args.iter().map(expr_to_code).collect();
            format!("{}({})", func.as_str(), arg_strs.join(", "))
        }
        Expression::MethodCall(obj, method, args) => {
            let arg_strs: Vec<String> = args.iter().map(expr_to_code).collect();
            format!("{}.{}({})", expr_to_code(obj), method.as_str(), arg_strs.join(", "))
        }
        Expression::New(type_name, args) => {
            let arg_strs: Vec<String> = args.iter().map(expr_to_code).collect();
            format!("New {}({})", type_name.as_str(), arg_strs.join(", "))
        }
        Expression::StringLiteral(s) => format!("\"{}\"", s.replace('"', "\"\"")),
        Expression::IntegerLiteral(i) => i.to_string(),
        Expression::BooleanLiteral(b) => if *b { "True".to_string() } else { "False".to_string() },
        Expression::DoubleLiteral(d) => d.to_string(),
        Expression::Me => "Me".to_string(),
        _ => "Nothing".to_string(),
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
            vybe_parser::MethodDecl::Sub(s)
                if s.name.as_str().eq_ignore_ascii_case("InitializeComponent") =>
            {
                Some(&s.body)
            }
            _ => None,
        }
    })?;

    let form_name = class_decl.name.as_str().to_string();
    let mut form = Form::new(&form_name);

    // builders keyed by field_name (the Me.X part), insertion_order preserves ordering
    let mut builders: std::collections::HashMap<String, ControlBuilder> =
        std::collections::HashMap::new();
    let mut insertion_order: Vec<String> = Vec::new();
    // parent_map: child_field_name -> parent_field_name
    let mut parent_map: std::collections::HashMap<String, String> =
        std::collections::HashMap::new();

    for stmt in init_method {
        match stmt {
            // ── Me.X = value  (form property or control registration) ──────────
            Statement::MemberAssignment { object, member, value } if is_me(object) => {
                let member_name = member.as_str();
                let member_lower = member_name.to_lowercase();

                match member_lower.as_str() {
                    "clientsize" => {
                        if let Some((w, h)) = extract_size(value) {
                            form.width = w;
                            form.height = h;
                        }
                    }
                    "text" => {
                        if let Expression::StringLiteral(s) = value {
                            form.text = s.clone();
                        }
                    }
                    "name" => {
                        if let Expression::StringLiteral(s) = value {
                            form.name = s.clone();
                        }
                    }
                    "backcolor" => {
                        if let Some(color) = extract_color(value) {
                            form.back_color = Some(color);
                        } else {
                            form.properties.set_raw("BackColor", expr_to_property_value(value));
                        }
                    }
                    "forecolor" => {
                        if let Some(color) = extract_color(value) {
                            form.fore_color = Some(color);
                        } else {
                            form.properties.set_raw("ForeColor", expr_to_property_value(value));
                        }
                    }
                    "font" => {
                        if let Some(font) = extract_font(value) {
                            form.font = Some(font);
                        } else {
                            form.properties.set_raw("Font", expr_to_property_value(value));
                        }
                    }
                    // Properties we deliberately ignore at parse time (purely designer
                    // metadata with no runtime effect in Vybe)
                    "autoscaledimensions" | "autoscalemode" | "padding" | "margin"
                    | "minimumsize" | "maximumsize" | "transparencykey"
                    | "topmost" | "opacity" => {}
                    _ => {
                        if let Expression::New(type_id, _) = value {
                            // Me.X = New SomeType() → register as a control
                            if let Some(ct) = vbnet_type_to_control_type(type_id.as_str()) {
                                if !builders.contains_key(member_name) {
                                    insertion_order.push(member_name.to_string());
                                }
                                builders.insert(
                                    member_name.to_string(),
                                    ControlBuilder::new(member_name.to_string(), ct),
                                );
                            }
                        } else {
                            // Unknown form-level property: store for round-trip fidelity
                            // (e.g. StartPosition, FormBorderStyle, MinimizeBox, etc.)
                            form.properties.set_raw(
                                member_name.to_string(),
                                expr_to_property_value(value),
                            );
                        }
                    }
                }
            }

            // ── Me.X.Prop = value  (control property assignment) ──────────────
            Statement::MemberAssignment { object, member, value } => {
                if let Some((ctrl_name, true)) = extract_me_member_target(object) {
                    let prop_name = member.as_str();
                    if let Some(builder) = builders.get_mut(&ctrl_name) {
                        apply_control_property(builder, prop_name, value);
                    }
                }
            }

            // ── ExpressionStatement method calls (DataBindings.Add, Controls.Add) ──
            Statement::ExpressionStatement(Expression::MethodCall(obj, method, args)) => {
                let method_lower = method.as_str().to_lowercase();

                if let Expression::MemberAccess(inner, member_id) = obj.as_ref() {
                    let member_str = member_id.as_str();

                    match method_lower.as_str() {
                        "add" if member_str.eq_ignore_ascii_case("DataBindings") => {
                            // Me.ctrl.DataBindings.Add("PropName", Me.bs, "Column" [, ...])
                            if let Some((ctrl_name, true)) = extract_me_member_target(inner) {
                                if args.len() >= 3 {
                                    let prop_name = match &args[0] {
                                        Expression::StringLiteral(s) => Some(s.clone()),
                                        _ => None,
                                    };
                                    let bs_name = match &args[1] {
                                        Expression::MemberAccess(inner2, m) if is_me(inner2) => {
                                            Some(m.as_str().to_string())
                                        }
                                        Expression::StringLiteral(s) => Some(s.clone()),
                                        _ => None,
                                    };
                                    let col_name = match &args[2] {
                                        Expression::StringLiteral(s) => Some(s.clone()),
                                        _ => None,
                                    };
                                    if let (Some(prop), Some(bs), Some(col)) =
                                        (prop_name, bs_name, col_name)
                                    {
                                        if let Some(builder) = builders.get_mut(&ctrl_name) {
                                            builder.extra_props.insert(
                                                "DataBindings.Source".to_string(),
                                                crate::properties::PropertyValue::String(bs),
                                            );
                                            builder.extra_props.insert(
                                                format!("DataBindings.{}", prop),
                                                crate::properties::PropertyValue::String(col),
                                            );
                                        }
                                    }
                                }
                            }
                        }
                        "add" if member_str.eq_ignore_ascii_case("Controls") => {
                            // Me.Container.Controls.Add(Me.Child) or Me.Controls.Add(Me.Child)
                            let parent_field = if is_me(inner) {
                                None // Form is parent
                            } else if let Some((name, true)) = extract_me_member_target(inner) {
                                Some(name)
                            } else {
                                continue;
                            };

                            if let Some(arg0) = args.first() {
                                if let Some((child_name, true)) =
                                    extract_me_member_target(arg0)
                                {
                                    if let Some(parent) = parent_field {
                                        parent_map.insert(child_name, parent);
                                    }
                                    // Form-as-parent: no entry needed (parent_id stays None)
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }

            // ── AddHandler Me.ctrl.Event, AddressOf Me.Handler ────────────
            // Wires events explicitly in InitializeComponent (alternative to Handles clause)
            Statement::AddHandler { event_target, handler } => {
                // Normalize: strip leading "Me." from both sides
                let target = if event_target.to_lowercase().starts_with("me.") {
                    &event_target[3..]
                } else {
                    event_target.as_str()
                };
                let handler_name = if handler.to_lowercase().starts_with("me.") {
                    handler[3..].to_string()
                } else {
                    handler.clone()
                };
                // Split on last '.' → (control_name, event_name)
                if let Some(dot_pos) = target.rfind('.') {
                    let ctrl_name = &target[..dot_pos];
                    let event_name = &target[dot_pos + 1..];
                    if let Some(event_type) = EventType::from_name(event_name) {
                        form.event_bindings.push(EventBinding::with_handler(
                            ctrl_name,
                            event_type,
                            handler_name,
                        ));
                    }
                }
            }

            // Skip other call statements (SuspendLayout, ResumeLayout, PerformLayout, etc.)
            // Also skip local variable declarations (e.g. Dim resources As ComponentResourceManager)
            Statement::ExpressionStatement(_) | Statement::Call { .. } => {}
            _ => {}
        }
    }

    // Build controls in insertion order, keeping field_name → Control pairing
    let mut built: Vec<(String, Control)> = Vec::new();
    for field_name in &insertion_order {
        if let Some(builder) = builders.remove(field_name) {
            built.push((field_name.clone(), builder.build()));
        }
    }

    // Build field_name → uuid map for parent resolution
    // Uses the original field names (e.g. "btn1_0") rather than ctrl.name (e.g. "btn1")
    let field_to_id: std::collections::HashMap<String, uuid::Uuid> =
        built.iter().map(|(n, c)| (n.clone(), c.id)).collect();

    // Apply parent_ids
    for (field_name, ctrl) in &mut built {
        if let Some(parent_field) = parent_map.get(field_name.as_str()) {
            if let Some(&parent_id) = field_to_id.get(parent_field.as_str()) {
                ctrl.parent_id = Some(parent_id);
            }
        }
    }

    for (_, ctrl) in built {
        form.controls.push(ctrl);
    }

    Some(form)
}

/// Apply a single property assignment (`prop_name = value`) to a ControlBuilder.
/// Handles well-known properties explicitly; everything else goes to extra_props for
/// round-trip fidelity (arbitrary designer properties like UseVisualStyleBackColor, etc.)
fn apply_control_property(
    builder: &mut ControlBuilder,
    prop_name: &str,
    value: &Expression,
) {
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
            } else {
                // Named/system color — preserve as Expression for round-trip
                builder.extra_props.insert(
                    "BackColor".to_string(),
                    expr_to_property_value(value),
                );
            }
        }
        "forecolor" => {
            if let Some(color) = extract_color(value) {
                builder.fore_color = Some(color);
            } else {
                builder.extra_props.insert(
                    "ForeColor".to_string(),
                    expr_to_property_value(value),
                );
            }
        }
        "font" => {
            if let Some(font) = extract_font(value) {
                builder.font = Some(font);
            } else {
                builder.extra_props.insert(
                    "Font".to_string(),
                    expr_to_property_value(value),
                );
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
            } else {
                builder.extra_props.insert("Tag".to_string(), expr_to_property_value(value));
            }
        }
        "tabindex" => {
            if let Expression::IntegerLiteral(n) = value {
                builder.tab_index = *n;
            }
        }
        // ── Data-source / binding properties ─────────────────────────────────
        "datasource" => {
            let pv = match value {
                Expression::MemberAccess(inner, m) if is_me(inner) => {
                    crate::properties::PropertyValue::String(m.as_str().to_string())
                }
                _ => expr_to_property_value(value),
            };
            builder.extra_props.insert("DataSource".to_string(), pv);
        }
        "datamember" => {
            builder.extra_props.insert("DataMember".to_string(), expr_to_property_value(value));
        }
        "displaymember" => {
            builder.extra_props.insert("DisplayMember".to_string(), expr_to_property_value(value));
        }
        "valuemember" => {
            builder.extra_props.insert("ValueMember".to_string(), expr_to_property_value(value));
        }
        "bindingsource" => {
            let pv = match value {
                Expression::MemberAccess(inner, m) if is_me(inner) => {
                    crate::properties::PropertyValue::String(m.as_str().to_string())
                }
                _ => expr_to_property_value(value),
            };
            builder.extra_props.insert("BindingSource".to_string(), pv);
        }
        "filter" => {
            builder.extra_props.insert("Filter".to_string(), expr_to_property_value(value));
        }
        "sort" => {
            builder.extra_props.insert("Sort".to_string(), expr_to_property_value(value));
        }
        "datasetname" => {
            builder.extra_props.insert("DataSetName".to_string(), expr_to_property_value(value));
        }
        "tablename" => {
            builder.extra_props.insert("TableName".to_string(), expr_to_property_value(value));
        }
        "selectcommand" => {
            builder.extra_props.insert("SelectCommand".to_string(), expr_to_property_value(value));
        }
        "connectionstring" => {
            builder.extra_props.insert("ConnectionString".to_string(), expr_to_property_value(value));
        }
        // ── Catch-all: arbitrary designer property ───────────────────────────
        // Preserves properties like UseVisualStyleBackColor, FlatStyle, Dock,
        // Anchor, Enabled, Visible, Multiline, ReadOnly, Minimum, Maximum,
        // DropDownStyle, AutoSize, Image, etc.
        _ => {
            builder.extra_props.insert(prop_name.to_string(), expr_to_property_value(value));
        }
    }
}

// ── Helpers ──────────────────────────────────────────────────────────────────

fn is_me(expr: &Expression) -> bool {
    matches!(expr, Expression::Me)
}

/// Given `MemberAccess(Me, "Button1")` → `("Button1", true)`.
fn extract_me_member_target(expr: &Expression) -> Option<(String, bool)> {
    if let Expression::MemberAccess(inner, member) = expr {
        if is_me(inner) {
            return Some((member.as_str().to_string(), true));
        }
    }
    None
}

/// Extract `(x, y)` from `New [System.Drawing.]Point(x, y)`.
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

/// Extract `(w, h)` from `New [System.Drawing.]Size(w, h)`.
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

/// Extract color as `#RRGGBB` from common VB.NET designer color expressions.
/// Returns `None` for named/system colors so the caller can fall back to Expression storage.
fn extract_color(expr: &Expression) -> Option<String> {
    match expr {
        Expression::Call(ident, args) => try_extract_color_call(ident.as_str(), args),
        Expression::MethodCall(_, ident, args) => try_extract_color_call(ident.as_str(), args),
        _ => None,
    }
}

fn try_extract_color_call(name: &str, args: &[Expression]) -> Option<String> {
    match name.to_lowercase().as_str() {
        "fromhtml" if args.len() == 1 => {
            if let Expression::StringLiteral(s) = &args[0] {
                return Some(s.clone());
            }
            None
        }
        "fromargb" => {
            if args.len() == 1 {
                let val = expr_to_i32(&args[0])? as u32;
                let r = ((val >> 16) & 0xFF) as u8;
                let g = ((val >> 8) & 0xFF) as u8;
                let b = (val & 0xFF) as u8;
                Some(format!("#{:02X}{:02X}{:02X}", r, g, b))
            } else if args.len() == 3 || args.len() == 4 {
                let start = if args.len() == 4 { 1 } else { 0 };
                let r = expr_to_i32(&args[start])? as u8;
                let g = expr_to_i32(&args[start + 1])? as u8;
                let b = expr_to_i32(&args[start + 2])? as u8;
                Some(format!("#{:02X}{:02X}{:02X}", r, g, b))
            } else {
                None
            }
        }
        _ => None,
    }
}

/// Extract font string `"Family, sizepx"` from `New [System.Drawing.]Font(family, size[, style])`.
/// The optional style argument is preserved via extra_props by the caller if needed.
fn extract_font(expr: &Expression) -> Option<String> {
    if let Expression::New(id, args) = expr {
        if last_component(id.as_str()).eq_ignore_ascii_case("Font") && args.len() >= 2 {
            if let Expression::StringLiteral(fam) = &args[0] {
                let size = match expr_to_i32(&args[1]) {
                    Some(i) => i as f32,
                    None => match &args[1] {
                        Expression::DoubleLiteral(d) => *d as f32,
                        _ => 12.0,
                    },
                };
                // If style arg present, append it so codegen can emit it back
                if args.len() >= 3 {
                    let style = expr_to_code(&args[2]);
                    return Some(format!("{}, {}px, {}", fam, size, style));
                }
                return Some(format!("{}, {}px", fam, size));
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
