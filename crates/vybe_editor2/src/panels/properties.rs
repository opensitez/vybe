use egui::Ui;
use vybe_forms::ControlType;
use crate::state::{EditorState, View};

#[derive(Clone, PartialEq)]
pub enum PropertiesTab { Properties, Events }

pub fn show(ui: &mut Ui, state: &mut EditorState, tab: &mut PropertiesTab) {
    let sel_id = state.selected_controls.first().copied();

    // Tab bar
    ui.horizontal(|ui| {
        if ui.selectable_label(*tab == PropertiesTab::Properties, "Properties").clicked() {
            *tab = PropertiesTab::Properties;
        }
        if ui.selectable_label(*tab == PropertiesTab::Events, "Events").clicked() {
            *tab = PropertiesTab::Events;
        }
    });
    ui.separator();

    match tab {
        PropertiesTab::Properties => show_properties(ui, state, sel_id),
        PropertiesTab::Events => show_events(ui, state, sel_id),
    }
}

fn show_properties(ui: &mut Ui, state: &mut EditorState, sel_id: Option<uuid::Uuid>) {
    if let Some(id) = sel_id {
        // Snapshot values before mutable borrow
        let snap = state.current_form_data().and_then(|f| f.controls.iter().find(|c| c.id == id).map(|c| (
            c.name.clone(),
            format!("{:?}", c.control_type),
            c.bounds.x, c.bounds.y, c.bounds.width, c.bounds.height,
            c.properties.get_string("Text").unwrap_or_default().to_string(),
            c.properties.get_string("BackColor").unwrap_or_default().to_string(),
            c.properties.get_string("ForeColor").unwrap_or_default().to_string(),
            c.properties.get_string("Font").unwrap_or_default().to_string(),
            c.properties.get_string("Enabled").map(|s| s != "false").unwrap_or(true),
            c.properties.get_string("Visible").map(|s| s != "false").unwrap_or(true),
            c.tab_index,
            c.control_type.clone(),
            c.properties.get_string("ConnectionString").unwrap_or_default().to_string(),
            c.properties.get_string("SelectCommand").unwrap_or_default().to_string(),
            c.properties.get_string("DataSource").unwrap_or_default().to_string(),
            c.properties.get_string("DataMember").unwrap_or_default().to_string(),
        )));

        let Some((
            mut name, ctype,
            x, y, w, h,
            mut text,
            mut back_color, mut fore_color, mut font,
            mut enabled, mut visible,
            tab_index,
            ct,
            mut conn_str, mut select_cmd,
            mut data_source, mut data_member,
        )) = snap else {
            ui.label("Control not found.");
            return;
        };

        let is_non_visual = ct.is_non_visual();

        egui::Grid::new("props").num_columns(2).striped(true).min_col_width(70.0).show(ui, |ui| {
            row_str(ui, "Name", &mut name);
            ui.end_row();
            ui.label("Type"); ui.label(egui::RichText::new(&ctype).weak()); ui.end_row();

            if !is_non_visual {
                let mut xi = x; row_int(ui, "Left", &mut xi); ui.end_row();
                let mut yi = y; row_int(ui, "Top", &mut yi); ui.end_row();
                let mut wi = w; row_int(ui, "Width", &mut wi); ui.end_row();
                let mut hi = h; row_int(ui, "Height", &mut hi); ui.end_row();

                // Apply geometry changes
                if xi != x || yi != y || wi != w || hi != h {
                    if let Some(form) = state.current_form_data_mut() {
                        if let Some(ctrl) = form.controls.iter_mut().find(|c| c.id == id) {
                            ctrl.bounds.x = xi; ctrl.bounds.y = yi;
                            ctrl.bounds.width = wi; ctrl.bounds.height = hi;
                        }
                    }
                }

                row_str(ui, "Text", &mut text); ui.end_row();
                row_str(ui, "BackColor", &mut back_color); ui.end_row();
                row_str(ui, "ForeColor", &mut fore_color); ui.end_row();
                row_str(ui, "Font", &mut font); ui.end_row();

                let mut ti = tab_index as i32;
                row_int(ui, "TabIndex", &mut ti); ui.end_row();

                ui.label("Enabled"); ui.checkbox(&mut enabled, ""); ui.end_row();
                ui.label("Visible"); ui.checkbox(&mut visible, ""); ui.end_row();
            }

            // Data binding properties
            match ct {
                ControlType::DataAdapterComponent => {
                    row_str(ui, "ConnectionString", &mut conn_str); ui.end_row();
                    row_str(ui, "SelectCommand", &mut select_cmd); ui.end_row();
                }
                ControlType::BindingSourceComponent => {
                    row_str(ui, "DataSource", &mut data_source); ui.end_row();
                    row_str(ui, "DataMember", &mut data_member); ui.end_row();
                }
                _ => {}
            }
        });

        // Write back all changes
        if let Some(form) = state.current_form_data_mut() {
            if let Some(ctrl) = form.controls.iter_mut().find(|c| c.id == id) {
                ctrl.name = name;
                if !is_non_visual {
                    ctrl.properties.set("Text".into(), text);
                    ctrl.properties.set("BackColor".into(), back_color);
                    ctrl.properties.set("ForeColor".into(), fore_color);
                    ctrl.properties.set("Font".into(), font);
                    ctrl.properties.set("Enabled".into(), enabled.to_string());
                    ctrl.properties.set("Visible".into(), visible.to_string());
                    ctrl.tab_index = tab_index;
                }
                match ctrl.control_type {
                    ControlType::DataAdapterComponent => {
                        ctrl.properties.set("ConnectionString".into(), conn_str);
                        ctrl.properties.set("SelectCommand".into(), select_cmd);
                    }
                    ControlType::BindingSourceComponent => {
                        ctrl.properties.set("DataSource".into(), data_source);
                        ctrl.properties.set("DataMember".into(), data_member);
                    }
                    _ => {}
                }
            }
        }
    } else {
        // Form properties
        let snap = state.current_form_data().map(|f| (
            f.name.clone(), f.text.clone(), f.width, f.height,
            f.back_color.clone().unwrap_or_default(),
        ));
        if let Some((fname, mut ftext, mut fw, mut fh, mut fbg)) = snap {
            ui.label(egui::RichText::new(format!("Form: {}", fname)).strong());
            ui.separator();
            egui::Grid::new("form_props").num_columns(2).striped(true).show(ui, |ui| {
                row_str(ui, "Text", &mut ftext); ui.end_row();
                row_int(ui, "Width", &mut fw); ui.end_row();
                row_int(ui, "Height", &mut fh); ui.end_row();
                row_str(ui, "BackColor", &mut fbg); ui.end_row();
            });
            if let Some(form) = state.current_form_data_mut() {
                form.text = ftext;
                form.width = fw;
                form.height = fh;
                if !fbg.is_empty() { form.back_color = Some(fbg); }
            }
        }
    }
}

fn show_events(ui: &mut Ui, state: &mut EditorState, sel_id: Option<uuid::Uuid>) {
    let (ctrl_name, form_name, ct) = if let Some(id) = sel_id {
        let snap = state.current_form_data().and_then(|f| {
            f.controls.iter().find(|c| c.id == id)
                .map(|c| (c.name.clone(), f.name.clone(), c.control_type.clone()))
        });
        match snap {
            Some((n, f, t)) => (n, f, Some(t)),
            None => return,
        }
    } else {
        let fname = state.current_form_data().map(|f| f.name.clone()).unwrap_or_default();
        (fname.clone(), fname, None)
    };

    let events = events_for_type(ct.as_ref());

    ui.label(egui::RichText::new(format!("{} events", ctrl_name)).strong());
    ui.separator();

    egui::ScrollArea::vertical().show(ui, |ui| {
        for event in events {
            if ui.selectable_label(false, event).clicked() {
                insert_event_handler(state, &form_name, &ctrl_name, event, ct.is_none());
            }
        }
    });
}

fn insert_event_handler(state: &mut EditorState, form_name: &str, ctrl_name: &str, event: &str, is_form: bool) {
    let handler_name = format!("{}_{}", ctrl_name, event);
    let e_type = event_args_type(event);
    let params = format!("sender As Object, e As {}", e_type);
    let handles = if is_form {
        format!("Handles Me.{}", event)
    } else {
        format!("Handles {}.{}", ctrl_name, event)
    };
    let stub = format!(
        "\n    Private Sub {}({}) {}\n        ' TODO\n    End Sub\n",
        handler_name, params, handles
    );

    let buf = state.get_code_buffer(form_name);
    if !buf.contains(&handler_name) {
        // Insert before End Class
        let lower = buf.to_lowercase();
        if let Some(idx) = lower.rfind("end class") {
            buf.insert_str(idx, &stub);
        } else {
            buf.push_str(&stub);
        }
    }
    state.view = View::CodeEditor;
    state.current_code_file = None; // use current_form for code buffer
}

// ── Helpers ───────────────────────────────────────────────────────────────────

fn row_str(ui: &mut Ui, label: &str, val: &mut String) {
    ui.label(label);
    ui.text_edit_singleline(val);
}

fn row_int<T: egui::emath::Numeric>(ui: &mut Ui, label: &str, val: &mut T) {
    ui.label(label);
    ui.add(egui::DragValue::new(val));
}

fn events_for_type(ct: Option<&ControlType>) -> &'static [&'static str] {
    match ct {
        None => &["Load", "Shown", "Activated", "Deactivate", "FormClosing", "FormClosed",
                  "Resize", "Paint", "Click", "DoubleClick", "KeyDown", "KeyUp", "KeyPress",
                  "MouseClick", "MouseDown", "MouseUp", "MouseMove"],
        Some(ControlType::Button) => &["Click", "MouseDown", "MouseUp", "MouseMove",
                  "MouseEnter", "MouseLeave", "GotFocus", "LostFocus", "KeyDown", "KeyUp", "KeyPress"],
        Some(ControlType::TextBox) | Some(ControlType::MaskedTextBox) =>
                  &["TextChanged", "KeyPress", "KeyDown", "KeyUp", "GotFocus", "LostFocus",
                    "Click", "Enter", "Leave", "Validating", "Validated"],
        Some(ControlType::Label) | Some(ControlType::LinkLabel) =>
                  &["Click", "DoubleClick", "MouseEnter", "MouseLeave"],
        Some(ControlType::CheckBox) | Some(ControlType::RadioButton) =>
                  &["CheckedChanged", "Click", "GotFocus", "LostFocus", "KeyPress"],
        Some(ControlType::ListBox) =>
                  &["SelectedIndexChanged", "Click", "DoubleClick", "GotFocus", "LostFocus", "KeyPress"],
        Some(ControlType::ComboBox) =>
                  &["SelectedIndexChanged", "SelectedValueChanged", "TextChanged",
                    "DropDown", "DropDownClosed", "Click", "GotFocus", "LostFocus", "KeyPress"],
        Some(ControlType::DataGridView) =>
                  &["CellClick", "CellDoubleClick", "CellValueChanged", "CellContentClick",
                    "CellEndEdit", "CellBeginEdit", "SelectionChanged", "RowEnter", "RowLeave",
                    "DataBindingComplete", "DataError", "KeyDown", "Scroll"],
        Some(ControlType::TreeView) =>
                  &["AfterSelect", "BeforeSelect", "AfterExpand", "AfterCollapse",
                    "NodeMouseClick", "NodeMouseDoubleClick", "AfterCheck", "KeyDown", "KeyPress"],
        Some(ControlType::ListView) =>
                  &["SelectedIndexChanged", "ItemActivate", "ColumnClick", "ItemCheck",
                    "Click", "DoubleClick", "KeyDown", "KeyPress"],
        Some(ControlType::TabControl) =>
                  &["SelectedIndexChanged", "Selected", "Click", "DoubleClick"],
        Some(ControlType::NumericUpDown) =>
                  &["ValueChanged", "KeyPress", "KeyDown", "GotFocus", "LostFocus"],
        Some(ControlType::TrackBar) =>
                  &["Scroll", "ValueChanged", "MouseDown", "MouseUp", "GotFocus", "LostFocus"],
        Some(ControlType::DateTimePicker) =>
                  &["ValueChanged", "DateChanged", "DropDown", "DropDownClosed", "GotFocus", "LostFocus"],
        Some(ControlType::BindingSourceComponent) =>
                  &["CurrentChanged", "PositionChanged", "DataSourceChanged"],
        Some(ControlType::Panel) | Some(ControlType::Frame) =>
                  &["Click", "DoubleClick", "MouseDown", "MouseUp", "MouseMove", "Paint", "Resize"],
        _ => &["Click", "DoubleClick", "MouseEnter", "MouseLeave"],
    }
}

fn event_args_type(event: &str) -> &'static str {
    match event.to_lowercase().as_str() {
        "mouseclick" | "mousedoubleclick" | "mousedown" | "mouseup" | "mousemove" | "mousewheel"
        | "nodemouseclick" | "nodemousedoubleclick" | "columnheadermouseclick" => "MouseEventArgs",
        "keydown" | "keyup" => "KeyEventArgs",
        "keypress" => "KeyPressEventArgs",
        "formclosing" => "FormClosingEventArgs",
        "formclosed" => "FormClosedEventArgs",
        "paint" | "cellpainting" => "PaintEventArgs",
        "cellclick" | "celldoubleclick" | "cellcontentclick" | "cellvaluechanged"
        | "cellendedit" | "cellbeginedit" | "cellvalidating" | "cellenter" | "cellleave"
        | "cellformatting" | "rowenter" | "rowleave" | "rowvalidating" | "rowvalidated" => "DataGridViewCellEventArgs",
        "dataerror" => "DataGridViewDataErrorEventArgs",
        "afterselect" | "beforeselect" | "afterexpand" | "aftercollapse"
        | "beforeexpand" | "beforecollapse" | "aftercheck" | "beforecheck" => "TreeViewEventArgs",
        "afterlabeledit" | "beforelabeledit" => "NodeLabelEditEventArgs",
        "linkclicked" => "LinkLabelLinkClickedEventArgs",
        "splittermoved" | "splittermoving" => "SplitterEventArgs",
        "scroll" => "ScrollEventArgs",
        "columnclick" => "ColumnClickEventArgs",
        "itemselectionchanged" => "ListViewItemSelectionChangedEventArgs",
        "navigating" => "WebBrowserNavigatingEventArgs",
        "navigated" => "WebBrowserNavigatedEventArgs",
        _ => "EventArgs",
    }
}
