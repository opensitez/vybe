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
        // Snapshot everything we need before any mutable borrow
        let snap = state.current_form_data().and_then(|f| {
            f.controls.iter().find(|c| c.id == id).map(|c| {
                let p = &c.properties;
                (
                    c.name.clone(),
                    format!("{}", c.control_type.as_str()),
                    c.bounds.x, c.bounds.y, c.bounds.width, c.bounds.height,
                    // Common appearance
                    p.get_string("Text").unwrap_or_default().to_string(),
                    p.get_string("BackColor").unwrap_or_default().to_string(),
                    p.get_string("ForeColor").unwrap_or_default().to_string(),
                    p.get_string("Font").unwrap_or_default().to_string(),
                    p.get_bool("Enabled").unwrap_or(true),
                    p.get_bool("Visible").unwrap_or(true),
                    c.tab_index,
                    c.control_type.clone(),
                    // Data binding (BindingSource)
                    p.get_string("DataSource").unwrap_or_default().to_string(),
                    p.get_string("DataMember").unwrap_or_default().to_string(),
                    p.get_string("Filter").unwrap_or_default().to_string(),
                    p.get_string("Sort").unwrap_or_default().to_string(),
                    // DataAdapter / SqlConnection
                    p.get_string("ConnectionString").unwrap_or_default().to_string(),
                    p.get_string("SelectCommand").unwrap_or_default().to_string(),
                    // DataSet
                    p.get_string("DataSetName").unwrap_or_default().to_string(),
                    // DataTable
                    p.get_string("TableName").unwrap_or_default().to_string(),
                    // Timer
                    p.get_string("Interval").unwrap_or_else(|| "100").to_string(),
                    // Grid / list binding
                    p.get_string("BindingSource").unwrap_or_default().to_string(),
                    // ComboBox / ListBox
                    p.get_string("DisplayMember").unwrap_or_default().to_string(),
                    p.get_string("ValueMember").unwrap_or_default().to_string(),
                    // Numeric/Range
                    p.get_string("Minimum").unwrap_or_else(|| "0").to_string(),
                    p.get_string("Maximum").unwrap_or_else(|| "100").to_string(),
                    // WebBrowser
                    p.get_string("URL").unwrap_or_default().to_string(),
                    // LinkLabel
                    p.get_string("LinkColor").unwrap_or_default().to_string(),
                    // CheckBox/RadioButton
                    p.get_bool("Checked").unwrap_or(false),
                    // RichTextBox
                    p.get_bool("ReadOnly").unwrap_or(false),
                    // TextBox
                    p.get_bool("Multiline").unwrap_or(false),
                    // Panel border style
                    p.get_string("BorderStyle").unwrap_or_default().to_string(),
                    // ProgressBar/TrackBar value
                    p.get_string("Value").unwrap_or_else(|| "0").to_string(),
                    // DataBindings — bindable properties for visual controls
                    p.get_string("DataBindings.Text").unwrap_or_default().to_string(),
                    p.get_string("DataBindings.Checked").unwrap_or_default().to_string(),
                    p.get_string("DataBindings.Value").unwrap_or_default().to_string(),
                    p.get_string("DataBindings.SelectedValue").unwrap_or_default().to_string(),
                    p.get_string("DataBindings.Visible").unwrap_or_default().to_string(),
                    p.get_string("DataBindings.Enabled").unwrap_or_default().to_string(),
                )
            })
        });

        let Some((
            mut name, ctype,
            x, y, w, h,
            mut text, mut back_color, mut fore_color, mut font,
            mut enabled, mut visible,
            tab_index,
            ct,
            mut data_source, mut data_member, mut filter, mut sort,
            mut conn_str, mut select_cmd,
            mut dataset_name,
            mut table_name,
            mut interval,
            mut binding_source,
            mut display_member, mut value_member,
            mut minimum, mut maximum,
            mut url,
            mut link_color,
            mut checked,
            mut read_only,
            mut multiline,
            mut border_style,
            mut value,
            mut db_text, mut db_checked, mut db_value,
            mut db_selected_value, mut db_visible, mut db_enabled,
        )) = snap else {
            ui.label("Control not found.");
            return;
        };

        let is_non_visual = ct.is_non_visual();

        // ── Control header ──────────────────────────────────────────────
        ui.label(egui::RichText::new(format!("{} ({})", name, ctype)).strong());
        ui.separator();

        egui::Grid::new("props")
            .num_columns(2)
            .striped(true)
            .min_col_width(80.0)
            .max_col_width(200.0)
            .show(ui, |ui| {

            // ── Identity ────────────────────────────────────────────────
            section_header(ui, "Identity");
            row_str(ui, "Name", &mut name); ui.end_row();

            // ── Layout (visual controls only) ───────────────────────────
            if !is_non_visual {
                section_header(ui, "Layout");
                let mut xi = x; row_int(ui, "Left",   &mut xi); ui.end_row();
                let mut yi = y; row_int(ui, "Top",    &mut yi); ui.end_row();
                let mut wi = w; row_int(ui, "Width",  &mut wi); ui.end_row();
                let mut hi = h; row_int(ui, "Height", &mut hi); ui.end_row();
                if xi != x || yi != y || wi != w || hi != h {
                    if let Some(form) = state.current_form_data_mut() {
                        if let Some(ctrl) = form.controls.iter_mut().find(|c| c.id == id) {
                            ctrl.bounds.x = xi; ctrl.bounds.y = yi;
                            ctrl.bounds.width = wi; ctrl.bounds.height = hi;
                        }
                    }
                }

                // ── Appearance ──────────────────────────────────────────
                section_header(ui, "Appearance");
                match ct {
                    ControlType::Label | ControlType::Button |
                    ControlType::CheckBox | ControlType::RadioButton |
                    ControlType::LinkLabel | ControlType::TextBox |
                    ControlType::RichTextBox | ControlType::MaskedTextBox |
                    ControlType::ListBox | ControlType::ComboBox |
                    ControlType::TabControl | ControlType::Frame |
                    ControlType::Panel => {
                        row_str(ui, "Text",      &mut text);      ui.end_row();
                    }
                    _ => {}
                }
                row_str(ui, "BackColor", &mut back_color); ui.end_row();
                row_str(ui, "ForeColor", &mut fore_color); ui.end_row();
                row_str(ui, "Font",      &mut font);       ui.end_row();

                // ── Behavior ────────────────────────────────────────────
                section_header(ui, "Behavior");
                let mut ti = tab_index as i32;
                row_int(ui, "TabIndex", &mut ti); ui.end_row();
                ui.label("Enabled"); ui.checkbox(&mut enabled, ""); ui.end_row();
                ui.label("Visible"); ui.checkbox(&mut visible, ""); ui.end_row();

                // ── Type-specific properties ────────────────────────────
                match ct {
                    ControlType::CheckBox | ControlType::RadioButton => {
                        section_header(ui, "State");
                        ui.label("Checked"); ui.checkbox(&mut checked, ""); ui.end_row();
                    }
                    ControlType::TextBox | ControlType::MaskedTextBox => {
                        section_header(ui, "Text");
                        ui.label("Multiline"); ui.checkbox(&mut multiline, ""); ui.end_row();
                    }
                    ControlType::RichTextBox => {
                        section_header(ui, "Text");
                        ui.label("ReadOnly"); ui.checkbox(&mut read_only, ""); ui.end_row();
                    }
                    ControlType::LinkLabel => {
                        section_header(ui, "Link");
                        row_str(ui, "LinkColor", &mut link_color); ui.end_row();
                    }
                    ControlType::Panel | ControlType::Frame => {
                        section_header(ui, "Appearance");
                        row_str(ui, "BorderStyle", &mut border_style); ui.end_row();
                    }
                    ControlType::ProgressBar | ControlType::TrackBar
                    | ControlType::NumericUpDown | ControlType::HScrollBar
                    | ControlType::VScrollBar => {
                        section_header(ui, "Range");
                        row_str(ui, "Minimum", &mut minimum); ui.end_row();
                        row_str(ui, "Maximum", &mut maximum); ui.end_row();
                        row_str(ui, "Value",   &mut value);   ui.end_row();
                    }
                    ControlType::WebBrowser => {
                        section_header(ui, "Navigation");
                        row_str(ui, "URL", &mut url); ui.end_row();
                    }
                    ControlType::DataGridView | ControlType::ListView
                    | ControlType::ListBox | ControlType::ComboBox
                    | ControlType::BindingNavigator => {
                        section_header(ui, "Data");
                        row_str(ui, "DataSource",     &mut data_source);    ui.end_row();
                        row_str(ui, "DataMember",     &mut data_member);    ui.end_row();
                        row_str(ui, "BindingSource",  &mut binding_source); ui.end_row();
                        if matches!(ct, ControlType::ComboBox | ControlType::ListBox) {
                            row_str(ui, "DisplayMember", &mut display_member); ui.end_row();
                            row_str(ui, "ValueMember",   &mut value_member);   ui.end_row();
                        }
                    }
                    _ => {}
                }
            }

            // ── Non-visual component properties ─────────────────────────
            match ct {
                ControlType::BindingSourceComponent => {
                    section_header(ui, "Data Binding");
                    row_str(ui, "DataSource", &mut data_source); ui.end_row();
                    row_str(ui, "DataMember", &mut data_member); ui.end_row();
                    row_str(ui, "Filter",     &mut filter);      ui.end_row();
                    row_str(ui, "Sort",       &mut sort);        ui.end_row();
                }
                ControlType::DataAdapterComponent => {
                    section_header(ui, "Database");
                    row_str(ui, "ConnectionString", &mut conn_str);    ui.end_row();
                    row_str(ui, "SelectCommand",    &mut select_cmd);  ui.end_row();
                }
                ControlType::SqlConnection | ControlType::OleDbConnection => {
                    section_header(ui, "Connection");
                    row_str(ui, "ConnectionString", &mut conn_str); ui.end_row();
                }
                ControlType::DataSetComponent => {
                    section_header(ui, "DataSet");
                    row_str(ui, "DataSetName", &mut dataset_name); ui.end_row();
                }
                ControlType::DataTableComponent => {
                    section_header(ui, "DataTable");
                    row_str(ui, "TableName", &mut table_name); ui.end_row();
                }
                ControlType::Timer => {
                    section_header(ui, "Timer");
                    row_str(ui, "Interval (ms)", &mut interval); ui.end_row();
                    ui.label("Enabled"); ui.checkbox(&mut enabled, ""); ui.end_row();
                }
                _ => {}
            }
        });

        // ── Write back all changes ───────────────────────────────────────
        if let Some(form) = state.current_form_data_mut() {
            if let Some(ctrl) = form.controls.iter_mut().find(|c| c.id == id) {
                ctrl.name = name;
                if !is_non_visual {
                    ctrl.properties.set("Text",      text);
                    ctrl.properties.set("BackColor",  back_color);
                    ctrl.properties.set("ForeColor",  fore_color);
                    ctrl.properties.set("Font",       font);
                    ctrl.properties.set("Enabled",    enabled.to_string());
                    ctrl.properties.set("Visible",    visible.to_string());
                    ctrl.properties.set("Checked",    checked.to_string());
                    ctrl.properties.set("ReadOnly",   read_only.to_string());
                    ctrl.properties.set("Multiline",  multiline.to_string());
                    ctrl.properties.set("BorderStyle", border_style);
                    ctrl.properties.set("Minimum",    minimum);
                    ctrl.properties.set("Maximum",    maximum);
                    ctrl.properties.set("Value",      value);
                    ctrl.properties.set("URL",        url);
                    ctrl.properties.set("LinkColor",  link_color);
                    ctrl.properties.set("DataSource",    data_source.clone());
                    ctrl.properties.set("DataMember",    data_member.clone());
                    ctrl.properties.set("BindingSource", binding_source);
                    ctrl.properties.set("DisplayMember", display_member);
                    ctrl.properties.set("ValueMember",   value_member);
                    ctrl.properties.set("DataBindings.Text",          db_text);
                    ctrl.properties.set("DataBindings.Checked",       db_checked);
                    ctrl.properties.set("DataBindings.Value",         db_value);
                    ctrl.properties.set("DataBindings.SelectedValue", db_selected_value);
                    ctrl.properties.set("DataBindings.Visible",       db_visible);
                    ctrl.properties.set("DataBindings.Enabled",       db_enabled);
                    ctrl.tab_index = tab_index;
                }
                match ctrl.control_type {
                    ControlType::BindingSourceComponent => {
                        ctrl.properties.set("DataSource", data_source);
                        ctrl.properties.set("DataMember", data_member);
                        ctrl.properties.set("Filter",     filter);
                        ctrl.properties.set("Sort",       sort);
                    }
                    ControlType::DataAdapterComponent => {
                        ctrl.properties.set("ConnectionString", conn_str);
                        ctrl.properties.set("SelectCommand",    select_cmd);
                    }
                    ControlType::SqlConnection | ControlType::OleDbConnection => {
                        ctrl.properties.set("ConnectionString", conn_str);
                    }
                    ControlType::DataSetComponent => {
                        ctrl.properties.set("DataSetName", dataset_name);
                    }
                    ControlType::DataTableComponent => {
                        ctrl.properties.set("TableName", table_name);
                    }
                    ControlType::Timer => {
                        ctrl.properties.set("Interval", interval);
                        ctrl.properties.set("Enabled",  enabled.to_string());
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
                row_str(ui, "Text",      &mut ftext); ui.end_row();
                row_int(ui, "Width",     &mut fw);    ui.end_row();
                row_int(ui, "Height",    &mut fh);    ui.end_row();
                row_str(ui, "BackColor", &mut fbg);   ui.end_row();
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
            if ui.selectable_label(false, *event).clicked() {
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
        let lower = buf.to_lowercase();
        if let Some(idx) = lower.rfind("end class") {
            buf.insert_str(idx, &stub);
        } else {
            buf.push_str(&stub);
        }
    }
    state.view = View::CodeEditor;
    state.current_code_file = None;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

fn section_header(ui: &mut Ui, label: &str) {
    ui.label(egui::RichText::new(label).small().weak().color(egui::Color32::from_rgb(100, 120, 160)));
    ui.end_row();
}

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
        Some(ControlType::Timer) => &["Tick"],
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
        "splittermoved" | "splittermoving" => "SplitterEventArgs",
        "scroll" => "ScrollEventArgs",
        "columnclick" => "ColumnClickEventArgs",
        _ => "EventArgs",
    }
}
