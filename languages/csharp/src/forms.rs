use std::sync::Arc;
use vybe_ast::*;
use vybe_platform_vybe::gui_state::GuiState;
use vybe_platform_dotnet::winforms::control::ControlType;
use vybe_platform_dotnet::winforms::form::Form;

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

/// Parse C# designer source and populate a `GuiState` with live widgets.
pub fn load_designer(source: &str, gui: &mut GuiState) -> Result<(), String> {
    let module = crate::parse(source)?;
    let init_body =
        find_initialize_component_in_module(&module).ok_or("no InitializeComponent found")?;

    let mut controls: Vec<ControlInfo> = Vec::new();

    for stmt in init_body {
        match &stmt.kind {
            StmtKind::Assign { targets, value } if targets.len() == 1 => {
                if let Some(field) = extract_this_field(&targets[0]) {
                    if let ExprKind::New { class, .. } = &value.kind {
                        let type_name = expr_to_type_name(class);
                        let short = last_component(&type_name);
                        if is_control_type(short) {
                            if !controls.iter().any(|c| c.name == field) {
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
                            }
                            continue;
                        }
                    }
                }

                if let Some((ctrl_name, event_name, handler_name)) =
                    extract_event_bind_from_assign(&targets[0], value)
                {
                    let normalized_event = event_name.to_ascii_lowercase();
                    gui.register_event(
                        ctrl_name,
                        &normalized_event,
                        vybe_bytecode::Value::String(Arc::from(handler_name)),
                    );
                    continue;
                }

                if let Some((ctrl_name, prop_name)) = extract_this_member_prop(&targets[0]) {
                    if let Some(c) = controls.iter_mut().find(|c| c.name == ctrl_name) {
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

            StmtKind::AddHandler {
                control,
                event,
                handler,
            } => {
                let ctrl_name = expr_to_control_name(control);
                let handler_name = expr_to_handler_name(handler);
                if !ctrl_name.is_empty() && !handler_name.is_empty() {
                    gui.register_event(
                        &ctrl_name,
                        event,
                        vybe_bytecode::Value::String(Arc::from(handler_name.as_str())),
                    );
                }
            }

            _ => {}
        }
    }

    for c in &controls {
        gui.add_widget(&c.type_name, &c.name, &c.text, c.x, c.y, c.width, c.height);
        for (prop, val) in &c.props {
            gui.set_property(&c.name, prop, val);
        }
    }

    Ok(())
}

/// Emit C# designer code from the current widget state.
pub fn save_designer(gui: &mut GuiState, class_name: &str) -> String {
    let mut out = String::new();
    out.push_str(&format!(
        "public partial class {} : System.Windows.Forms.Form\n{{\n",
        class_name
    ));

    struct CtrlSnapshot {
        name: String,
        type_name: &'static str,
        x: i32,
        y: i32,
        w: i32,
        h: i32,
    }

    let mut snaps: Vec<CtrlSnapshot> = Vec::new();
    for i in 0..gui.form.control_count() {
        if let Some(ctrl) = gui.form.control(i) {
            let name = ctrl.name();
            if name.is_empty() {
                continue;
            }
            let r = ctrl.rect();
            snaps.push(CtrlSnapshot {
                name: name.to_string(),
                type_name: widget_to_csharp_type(ctrl),
                x: r.x as i32,
                y: r.y as i32,
                w: r.w as i32,
                h: r.h as i32,
            });
        }
    }

    for s in &snaps {
        out.push_str(&format!("    private {} {};\n", s.type_name, s.name));
    }

    out.push_str("\n    private void InitializeComponent()\n    {\n");
    out.push_str("        this.SuspendLayout();\n");

    for s in &snaps {
        out.push_str(&format!(
            "        this.{0} = new {1}();\n",
            s.name, s.type_name
        ));
        out.push_str(&format!(
            "        this.{0}.Location = new System.Drawing.Point({1}, {2});\n",
            s.name, s.x, s.y
        ));
        out.push_str(&format!(
            "        this.{0}.Size = new System.Drawing.Size({1}, {2});\n",
            s.name, s.w, s.h
        ));
        out.push_str(&format!("        this.{0}.Name = \"{0}\";\n", s.name));

        let text = gui.get_property(&s.name, "text");
        if !text.is_empty() {
            out.push_str(&format!(
                "        this.{0}.Text = \"{1}\";\n",
                s.name,
                text.replace('"', "\\\"")
            ));
        }

        out.push_str(&format!("        this.Controls.Add(this.{});\n", s.name));
    }

    out.push_str(&format!(
        "        this.ClientSize = new System.Drawing.Size({}, {});\n",
        gui.width, gui.height
    ));
    out.push_str(&format!(
        "        this.Text = \"{}\";\n",
        gui.form.title.replace('"', "\\\"")
    ));
    out.push_str("        this.ResumeLayout(false);\n");
    out.push_str("    }\n}\n");
    out
}

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

fn find_initialize_component<'a>(class: &'a StmtKind) -> Option<&'a Vec<Statement>> {
    if let StmtKind::ClassDecl { members, .. } = class {
        for member in members {
            if let ClassMember::Method(stmt) = member {
                if let StmtKind::FunctionDecl { name, body, .. } = &stmt.kind {
                    if name.eq_ignore_ascii_case("InitializeComponent") {
                        return Some(body);
                    }
                }
            }
        }
    }
    None
}

fn find_initialize_component_in_module<'a>(module: &'a Module) -> Option<&'a Vec<Statement>> {
    for stmt in &module.body {
        if let Some(body) = find_initialize_component(&stmt.kind) {
            return Some(body);
        }
    }
    None
}

fn extract_this_field<'a>(expr: &'a Expression) -> Option<&'a str> {
    if let ExprKind::Member { object, field, .. } = &expr.kind {
        if matches!(object.kind, ExprKind::This) {
            return Some(field.as_str());
        }
    }
    None
}

fn extract_this_member_prop<'a>(expr: &'a Expression) -> Option<(&'a str, &'a str)> {
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

fn extract_event_bind_from_assign<'a>(
    target: &'a Expression,
    value: &'a Expression,
) -> Option<(&'a str, &'a str, &'a str)> {
    let (ctrl, event_name) = extract_this_member_prop(target)?;

    if let ExprKind::Binary {
        op: BinOp::Add,
        left,
        right,
    } = &value.kind
    {
        let (left_ctrl, left_event) = extract_this_member_prop(left)?;
        if left_ctrl == ctrl && left_event == event_name {
            if let ExprKind::Member { field, .. } = &right.kind {
                return Some((ctrl, event_name, field.as_str()));
            }
            if let ExprKind::Ident(name) = &right.kind {
                return Some((ctrl, event_name, name.as_str()));
            }
        }
    }

    None
}

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

fn last_component(name: &str) -> &str {
    name.rsplit('.').next().unwrap_or(name)
}

fn expr_to_control_name(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Ident(s) => s.clone(),
        ExprKind::Member { field, .. } => field.clone(),
        _ => String::new(),
    }
}

fn expr_to_handler_name(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Ident(s) => s.clone(),
        ExprKind::Member { field, .. } => field.clone(),
        _ => String::new(),
    }
}

fn expr_to_value_string(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(s)) => s.clone(),
        ExprKind::Lit(Literal::Int(n)) => n.to_string(),
        ExprKind::Lit(Literal::Float(f)) => f.to_string(),
        ExprKind::Lit(Literal::Bool(b)) => b.to_string(),
        ExprKind::Lit(Literal::Null) => "null".to_string(),
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
            format!("new {}({})", name, arg_strs.join(", "))
        }
        ExprKind::Call { callee, args, .. } => {
            let name = expr_to_value_string(callee);
            let arg_strs: Vec<String> = args
                .iter()
                .map(|a| expr_to_value_string(&a.value))
                .collect();
            format!("{}({})", name, arg_strs.join(", "))
        }
        _ => format!("{:?}", expr.kind),
    }
}

fn widget_to_csharp_type(widget: &dyn vybe_widgets::PanelWidget) -> &'static str {
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
