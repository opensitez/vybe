//! VB.NET designer form loading and saving.
//!
//! **Load**: parse `.Designer.vb` with `vb::parse()` → walk `InitializeComponent`
//! in the AST → populate a `GuiState` with real `vybe_widgets` directly.
//!
//! **Save**: walk the live widgets on the `GuiState` → emit VB.NET designer code.
//!
//! No intermediate data model — the designer works with the same widgets the
//! runtime uses, giving true WYSIWYG.

use vybe_ast::*;
use vybe_platform_vybe::gui_state::GuiState;
use vybe_platform_dotnet::winforms::control::ControlType;

// ═══════════════════════════════════════════════════════════════════════════
// Load: AST → GuiState
// ═══════════════════════════════════════════════════════════════════════════

/// Parse VB designer source and populate a `GuiState` with live widgets.
///
/// Parses the combined designer + user code, finds the class, locates
/// `InitializeComponent`, and walks its statements to create widgets,
/// set properties, and wire events — all through `GuiState`'s existing API.
pub fn load_designer(source: &str, gui: &mut GuiState) -> Result<(), String> {
    let module = super::parse(source)?;
    let class = find_class(&module).ok_or("no class found in designer source")?;
    let init_body = find_initialize_component(&class).ok_or("no InitializeComponent found")?;

    // Collect Handles clauses from methods for event wiring.
    let handles = collect_handles(&class);

    // Two-pass approach:
    // Pass 1: collect control registrations (Me.X = New Type()) and their properties.
    // Pass 2: add widgets to GuiState in the right order, then apply properties.

    let mut controls: Vec<ControlInfo> = Vec::new();
    let mut form_props: Vec<(&str, &Expression)> = Vec::new();

    for stmt in init_body {
        match &stmt.kind {
            // Me.X = New SomeType(args) — register a control
            StmtKind::Assign { targets, value } if targets.len() == 1 => {
                if let Some(field) = extract_me_field(&targets[0]) {
                    if let ExprKind::New { class: cls, .. } = &value.kind {
                        let type_name = expr_to_type_name(cls);
                        let short = last_component(&type_name);
                        if is_control_type(short) {
                            controls.push(ControlInfo {
                                name: field.to_string(),
                                type_name: short.to_string(),
                                x: 0,
                                y: 0,
                                width: 100,
                                height: 30,
                                text: String::new(),
                                props: Vec::new(),
                            });
                            continue;
                        }
                    }
                    // Me.ClientSize, Me.Text, etc. — form-level property
                    form_props.push((field, value));
                }
                // Me.X.Prop = value — control property
                if let Some((ctrl_name, prop_name)) = extract_me_member_prop(&targets[0]) {
                    if let Some(c) = controls
                        .iter_mut()
                        .find(|c| c.name.eq_ignore_ascii_case(ctrl_name))
                    {
                        if prop_name.eq_ignore_ascii_case("Location") {
                            if let Some((x, y)) = extract_point(value) {
                                c.x = x;
                                c.y = y;
                            }
                        } else if prop_name.eq_ignore_ascii_case("Size") {
                            if let Some((w, h)) = extract_size(value) {
                                c.width = w;
                                c.height = h;
                            }
                        } else if prop_name.eq_ignore_ascii_case("ClientSize") {
                            if let Some((w, h)) = extract_size(value) {
                                c.width = w;
                                c.height = h;
                            }
                        } else if prop_name.eq_ignore_ascii_case("Text") {
                            if let Some(s) = extract_string(value) {
                                c.text = s;
                            }
                        } else {
                            c.props
                                .push((prop_name.to_string(), expr_to_value_string(value)));
                        }
                    }
                }
            }

            // AddHandler control.Event, AddressOf handler
            StmtKind::AddHandler {
                control,
                event,
                handler,
            } => {
                let ctrl_name = expr_to_control_name(control);
                let handler_name = expr_to_handler_name(handler);
                gui.register_event(
                    &ctrl_name,
                    event,
                    vybe_bytecode::Value::String(std::sync::Arc::from(handler_name.as_str())),
                );
            }

            _ => {}
        }
    }

    // Apply form-level properties
    for (prop, value) in &form_props {
        match prop.to_lowercase().as_str() {
            "clientsize" => {
                if let Some((w, h)) = extract_size(value) {
                    gui.width = w as u32;
                    gui.height = h as u32;
                }
            }
            "text" => {
                if let Some(s) = extract_string(value) {
                    gui.form.title = s;
                }
            }
            _ => {
                gui.set_property("form", prop, &expr_to_value_string(value));
            }
        }
    }

    // Add controls to GuiState
    for c in &controls {
        gui.add_widget(&c.type_name, &c.name, &c.text, c.x, c.y, c.width, c.height);
        for (prop, val) in &c.props {
            gui.set_property(&c.name, prop, val);
        }
    }

    // Wire Handles clauses as events
    for (handler_name, handle_targets) in &handles {
        for target in handle_targets {
            // target is e.g. "btn1.Click"
            if let Some(dot) = target.rfind('.') {
                let ctrl = &target[..dot];
                let event = &target[dot + 1..];
                // Normalize: strip "Me." or "MyBase." prefix
                let ctrl = ctrl
                    .strip_prefix("Me.")
                    .or_else(|| ctrl.strip_prefix("MyBase."))
                    .unwrap_or(ctrl);
                gui.register_event(
                    ctrl,
                    event,
                    vybe_bytecode::Value::String(std::sync::Arc::from(handler_name.as_str())),
                );
            }
        }
    }

    Ok(())
}

// ═══════════════════════════════════════════════════════════════════════════
// Save: GuiState → VB designer code
// ═══════════════════════════════════════════════════════════════════════════

/// Emit VB.NET designer code from the current widget state.
pub fn save_designer(gui: &mut GuiState, class_name: &str) -> String {
    let mut out = String::new();
    out.push_str(&format!(
        "<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _\n"
    ));
    out.push_str(&format!("Partial Class {}\n", class_name));
    out.push_str("    Inherits System.Windows.Forms.Form\n\n");

    // Field declarations
    for i in 0..gui.form.control_count() {
        if let Some(ctrl) = gui.form.control(i) {
            let name = ctrl.name();
            if !name.is_empty() {
                let type_name = widget_to_vb_type(ctrl);
                out.push_str(&format!(
                    "    Friend WithEvents {} As {}\n",
                    name, type_name
                ));
            }
        }
    }

    out.push_str("\n    Private Sub InitializeComponent()\n");

    out.push_str("        Me.SuspendLayout()\n");

    // Collect control geometry from widgets (immutable borrow of form)
    struct CtrlSnapshot {
        name: String,
        type_name: &'static str,
        x: i32,
        y: i32,
        w: i32,
        h: i32,
    }
    let mut snapshots: Vec<CtrlSnapshot> = Vec::new();
    for i in 0..gui.form.control_count() {
        if let Some(ctrl) = gui.form.control(i) {
            let name = ctrl.name();
            if name.is_empty() {
                continue;
            }
            let r = ctrl.rect();
            snapshots.push(CtrlSnapshot {
                name: name.to_string(),
                type_name: widget_to_vb_type(ctrl),
                x: r.x as i32,
                y: r.y as i32,
                w: r.w as i32,
                h: r.h as i32,
            });
        }
    }

    // Control instantiation
    for snap in &snapshots {
        out.push_str(&format!(
            "        Me.{} = New {}()\n",
            snap.name, snap.type_name
        ));
    }

    // Emit control properties (can now freely borrow gui)
    for snap in &snapshots {
        out.push_str(&format!(
            "        Me.{}.Location = New System.Drawing.Point({}, {})\n",
            snap.name, snap.x, snap.y
        ));
        out.push_str(&format!(
            "        Me.{}.Size = New System.Drawing.Size({}, {})\n",
            snap.name, snap.w, snap.h
        ));
        out.push_str(&format!(
            "        Me.{}.Name = \"{}\"\n",
            snap.name, snap.name
        ));

        // Text from widget
        let text = gui.get_property(&snap.name, "text");
        if !text.is_empty() {
            out.push_str(&format!(
                "        Me.{}.Text = \"{}\"\n",
                snap.name,
                text.replace('"', "\"\"")
            ));
        }

        // Extra properties from the property store
        let name_lower = snap.name.to_lowercase();
        let keys: Vec<(String, String)> = gui
            .properties
            .keys()
            .filter(|(n, _)| *n == name_lower)
            .cloned()
            .collect();
        for (_, prop) in keys {
            if matches!(prop.as_str(), "text" | "name") {
                continue;
            }
            if let Some(val) = gui.properties.get(&(name_lower.clone(), prop.clone())) {
                out.push_str(&format!(
                    "        Me.{}.{} = {}\n",
                    snap.name,
                    capitalize(&prop),
                    val
                ));
            }
        }
    }

    // Controls.Add
    for snap in &snapshots {
        out.push_str(&format!("        Me.Controls.Add(Me.{})\n", snap.name));
    }

    // Form properties
    out.push_str(&format!(
        "        Me.ClientSize = New System.Drawing.Size({}, {})\n",
        gui.width, gui.height
    ));
    if !gui.form.title.is_empty() {
        out.push_str(&format!(
            "        Me.Text = \"{}\"\n",
            gui.form.title.replace('"', "\"\"")
        ));
    }
    out.push_str(&format!("        Me.Name = \"{}\"\n", class_name));
    out.push_str("        Me.ResumeLayout(False)\n");
    out.push_str("        Me.PerformLayout()\n");
    out.push_str("    End Sub\n\n");
    out.push_str("End Class\n");

    out
}

// ═══════════════════════════════════════════════════════════════════════════
// Helpers
// ═══════════════════════════════════════════════════════════════════════════

struct ControlInfo {
    name: String,
    type_name: String,
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    text: String,
    props: Vec<(String, String)>,
}

/// Find the first ClassDecl in the module.
fn find_class(module: &Module) -> Option<&StmtKind> {
    for stmt in &module.body {
        if let StmtKind::ClassDecl { .. } = &stmt.kind {
            return Some(&stmt.kind);
        }
    }
    None
}

/// Find the body of `InitializeComponent` in a class.
fn find_initialize_component<'a>(class: &'a StmtKind) -> Option<&'a Vec<Statement>> {
    if let StmtKind::ClassDecl { members, .. } = class {
        for member in members {
            if let ClassMember::Method(stmt) = member {
                if let StmtKind::FunctionDecl {
                    name,
                    body,
                    is_sub: true,
                    ..
                } = &stmt.kind
                {
                    if name.eq_ignore_ascii_case("InitializeComponent") {
                        return Some(body);
                    }
                }
            }
        }
    }
    None
}

/// Collect all Handles clauses: method_name → vec of "control.event" strings.
fn collect_handles(class: &StmtKind) -> Vec<(String, Vec<String>)> {
    let mut result = Vec::new();
    if let StmtKind::ClassDecl { members, .. } = class {
        for member in members {
            if let ClassMember::Method(stmt) = member {
                if let StmtKind::FunctionDecl { name, handles, .. } = &stmt.kind {
                    if !handles.is_empty() {
                        result.push((name.clone(), handles.clone()));
                    }
                }
            }
        }
    }
    result
}

/// `Me.X` → Some("X")
fn extract_me_field<'a>(expr: &'a Expression) -> Option<&'a str> {
    if let ExprKind::Member { object, field, .. } = &expr.kind {
        if matches!(object.kind, ExprKind::This) {
            return Some(field.as_str());
        }
    }
    None
}

/// `Me.X.Prop` → Some(("X", "Prop"))
fn extract_me_member_prop<'a>(expr: &'a Expression) -> Option<(&'a str, &'a str)> {
    if let ExprKind::Member {
        object,
        field: prop,
        ..
    } = &expr.kind
    {
        if let ExprKind::Member {
            object: inner,
            field: ctrl,
            ..
        } = &object.kind
        {
            if matches!(inner.kind, ExprKind::This) {
                return Some((ctrl.as_str(), prop.as_str()));
            }
        }
    }
    None
}

/// `New System.Drawing.Point(x, y)` → Some((x, y))
fn extract_point(expr: &Expression) -> Option<(i32, i32)> {
    if let ExprKind::New { class, args } = &expr.kind {
        let name = expr_to_type_name(class);
        if last_component(&name).eq_ignore_ascii_case("Point") && args.len() == 2 {
            let x = arg_to_i32(&args[0])?;
            let y = arg_to_i32(&args[1])?;
            return Some((x, y));
        }
    }
    None
}

/// `New System.Drawing.Size(w, h)` → Some((w, h))
fn extract_size(expr: &Expression) -> Option<(i32, i32)> {
    if let ExprKind::New { class, args } = &expr.kind {
        let name = expr_to_type_name(class);
        if last_component(&name).eq_ignore_ascii_case("Size") && args.len() == 2 {
            let w = arg_to_i32(&args[0])?;
            let h = arg_to_i32(&args[1])?;
            return Some((w, h));
        }
    }
    None
}

fn extract_string(expr: &Expression) -> Option<String> {
    if let ExprKind::Lit(Literal::Str(s)) = &expr.kind {
        Some(s.clone())
    } else {
        None
    }
}

fn arg_to_i32(arg: &Argument) -> Option<i32> {
    match &arg.value.kind {
        ExprKind::Lit(Literal::Int(n)) => Some(*n as i32),
        ExprKind::Lit(Literal::Float(f)) => Some(*f as i32),
        _ => None,
    }
}

/// Flatten an expression to a type name string (e.g. `System.Windows.Forms.Button` → "System.Windows.Forms.Button").
fn expr_to_type_name(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Ident(s) => s.clone(),
        ExprKind::Member { object, field, .. } => {
            let base = expr_to_type_name(object);
            format!("{}.{}", base, field)
        }
        _ => String::new(),
    }
}

/// `"System.Windows.Forms.Button"` → `"Button"`
fn last_component(name: &str) -> &str {
    name.rsplit('.').next().unwrap_or(name)
}

/// Extract control name from an AddHandler control expression.
fn expr_to_control_name(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Ident(s) => s.to_lowercase(),
        ExprKind::Member { object, field, .. } => {
            if matches!(object.kind, ExprKind::This) {
                field.to_lowercase()
            } else {
                field.to_lowercase()
            }
        }
        _ => String::new(),
    }
}

/// Extract handler name from an AddHandler handler expression.
fn expr_to_handler_name(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Ident(s) => s.clone(),
        ExprKind::Member { field, .. } => field.clone(),
        ExprKind::AddressOf(s) => s.clone(),
        _ => String::new(),
    }
}

/// Convert an expression to a string value for `set_property`.
fn expr_to_value_string(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(s)) => s.clone(),
        ExprKind::Lit(Literal::Int(n)) => n.to_string(),
        ExprKind::Lit(Literal::Float(f)) => f.to_string(),
        ExprKind::Lit(Literal::Bool(b)) => {
            if *b {
                "True".to_string()
            } else {
                "False".to_string()
            }
        }
        ExprKind::Lit(Literal::Null) => "Nothing".to_string(),
        ExprKind::Ident(s) => s.clone(),
        ExprKind::Member { object, field, .. } => {
            format!("{}.{}", expr_to_value_string(object), field)
        }
        ExprKind::New { class, args } => {
            let name = expr_to_type_name(class);
            let arg_strs: Vec<String> = args
                .iter()
                .map(|a| expr_to_value_string(&a.value))
                .collect();
            format!("New {}({})", name, arg_strs.join(", "))
        }
        ExprKind::Call { callee, args, .. } => {
            let name = expr_to_value_string(callee);
            let arg_strs: Vec<String> = args
                .iter()
                .map(|a| expr_to_value_string(&a.value))
                .collect();
            format!("{}({})", name, arg_strs.join(", "))
        }
        ExprKind::This => "Me".to_string(),
        _ => format!("{:?}", expr.kind),
    }
}

/// Map a widget type back to a VB.NET type name for codegen.
fn widget_to_vb_type(widget: &dyn vybe_widgets::PanelWidget) -> &'static str {
    ControlType::dotnet_class_name_for_widget_type_name(std::any::type_name_of_val(widget))
}

fn is_control_type(name: &str) -> bool {
    matches!(
        name.to_lowercase().as_str(),
        "button"
            | "label"
            | "textbox"
            | "checkbox"
            | "radiobutton"
            | "combobox"
            | "listbox"
            | "panel"
            | "groupbox"
            | "picturebox"
            | "progressbar"
            | "trackbar"
            | "numericupdown"
            | "datetimepicker"
            | "richtextbox"
            | "treeview"
            | "datagridview"
            | "datagrid"
            | "listview"
            | "tabcontrol"
            | "tabpage"
            | "monthcalendar"
            | "hscrollbar"
            | "vscrollbar"
            | "menustrip"
            | "toolstrip"
            | "statusstrip"
            | "contextmenustrip"
            | "splitcontainer"
            | "flowlayoutpanel"
            | "tablelayoutpanel"
            | "maskedtextbox"
            | "linklabel"
            | "bindingnavigator"
            | "usercontrol"
            | "webbrowser"
            | "canvas"
            | "paintbox"
            | "timer"
            | "imagelist"
            | "tooltip"
            | "errorprovider"
            | "openfiledialog"
            | "savefiledialog"
            | "fontdialog"
            | "colordialog"
            | "folderbrowserdialog"
            | "printdialog"
            | "notifyicon"
            | "helpprovider"
            | "backgroundworker"
            | "bindingsource"
            | "dataset"
            | "datatable"
            | "dataadapter"
    )
}

fn capitalize(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        None => String::new(),
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
    }
}
