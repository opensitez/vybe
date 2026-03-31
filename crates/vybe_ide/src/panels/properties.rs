use egui::Ui;
use vybe_forms::ControlType;
use crate::state::{EditorState, View};

#[derive(Clone, PartialEq)]
pub enum PropertiesTab { Properties, Events }

// Internal temporary state used by properties panel across frames
#[derive(Default)]
pub struct PropertiesState {
    pub show_conn_builder: bool,
    pub conn_status: String,
    pub da_tables: Vec<String>,
}

pub fn show(ui: &mut Ui, state: &mut EditorState, tab: &mut PropertiesTab) {
    let sel_id = state.selected_controls.first().copied();

    // Give ui.memory an ID to persist our local builder state
    let state_id = ui.id().with("properties_internal_state");
    let local_state = ui.memory_mut(|mem| {
        if let Some(arc) = mem.data.get_temp::<std::sync::Arc<std::sync::Mutex<PropertiesState>>>(state_id) {
            arc
        } else {
            let new_state = std::sync::Arc::new(std::sync::Mutex::new(PropertiesState::default()));
            mem.data.insert_temp(state_id, new_state.clone());
            new_state
        }
    });

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
        PropertiesTab::Properties => show_properties(ui, state, sel_id, &local_state),
        PropertiesTab::Events => show_events(ui, state, sel_id),
    }
}

fn show_properties(ui: &mut Ui, state: &mut EditorState, sel_id: Option<uuid::Uuid>, local_state_arc: &std::sync::Arc<std::sync::Mutex<PropertiesState>>) {
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
                    p.get_string("DbType").unwrap_or("SQLite").to_string(),
                    p.get_string("DbPath").unwrap_or_default().to_string(),
                    p.get_string("DbHost").unwrap_or("localhost").to_string(),
                    p.get_string("DbPort").unwrap_or_default().to_string(),
                    p.get_string("DbName").unwrap_or_default().to_string(),
                    p.get_string("DbUser").unwrap_or_default().to_string(),
                    p.get_string("DbPassword").unwrap_or_default().to_string(),
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
                    p.get_string("DataBindings.Source").unwrap_or_default().to_string(),
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
            mut db_type, mut db_path, mut db_host, mut db_port, mut db_name, mut db_user, mut db_pass,
            mut dataset_name,
            mut table_name,
            mut interval,
            _binding_source,
            mut display_member, mut value_member,
            mut minimum, mut maximum,
            mut url,
            mut link_color,
            mut checked,
            mut read_only,
            mut multiline,
            mut border_style,
            mut value,
            mut db_source, mut db_text, mut db_checked, mut db_value,
            db_selected_value, db_visible, db_enabled,
        )) = snap else {
            ui.label("Control not found.");
            return;
        };

        // Snapshot all data-related components on the form for combo boxes
        let mut form_binding_sources = Vec::new();
        let mut form_data_sources = Vec::new();
        if let Some(f) = state.current_form_data() {
            for c in &f.controls {
                if c.id == id { continue; }
                if matches!(c.control_type, ControlType::BindingSourceComponent) {
                    form_binding_sources.push(c.name.clone());
                }
                if matches!(c.control_type, ControlType::DataAdapterComponent | ControlType::DataSetComponent | ControlType::DataTableComponent) {
                    form_data_sources.push(c.name.clone());
                }
            }
        }

        let is_non_visual = ct.is_non_visual();

        // ── Helper: Resolve absolute database path ──────────────────────
        let project_dir = state.project_path.clone().and_then(|p| p.parent().map(|d| d.to_path_buf()))
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_else(|_| std::path::PathBuf::from(".")));
            
        let resolve_conn_str = |conn_str_in: &str| -> String {
            let lower = conn_str_in.to_lowercase();
            if let Some(pos) = lower.find("data source=") {
                let start = pos + 12;
                let rest = &conn_str_in[start..];
                let end = rest.find(';').unwrap_or(rest.len());
                let db_path = rest[..end].trim();
                if !db_path.is_empty() && db_path != ":memory:" && !std::path::Path::new(db_path).is_absolute() {
                    let abs = project_dir.join(db_path);
                    return format!("{}Data Source={}{}", &conn_str_in[..pos], abs.display(), &rest[end..]);
                }
            }
            conn_str_in.to_string()
        };

        // ── Helper: Resolve columns for a generic BindingSource name ─────
        let resolve_columns_for_bs = |state: &EditorState, bs_name: &str, rc_str: &dyn Fn(&str) -> String| -> Vec<String> {
            if bs_name.is_empty() { return Vec::new(); }
            let f = match state.current_form_data() { Some(f) => f, None => return Vec::new() };
            let bs_ctrl = match f.controls.iter().find(|c| c.name.eq_ignore_ascii_case(bs_name) && matches!(c.control_type, ControlType::BindingSourceComponent)) {
                Some(c) => c, None => return Vec::new()
            };
            let da_name = bs_ctrl.properties.get_string("DataSource").unwrap_or_default();
            let data_member = bs_ctrl.properties.get_string("DataMember").unwrap_or_default();
            
            let da_ctrl = match f.controls.iter().find(|c| c.name.eq_ignore_ascii_case(da_name) && matches!(c.control_type, ControlType::DataAdapterComponent)) {
                Some(c) => c, None => return Vec::new()
            };
            let da_conn = da_ctrl.properties.get_string("ConnectionString").unwrap_or_default();
            if da_conn.is_empty() { return Vec::new(); }
            
            let query = da_ctrl.properties.get_string("SelectCommand").unwrap_or_default();
            let final_query = if !query.is_empty() { query.to_string() } else if !data_member.is_empty() { format!("SELECT * FROM {}", data_member) } else { return Vec::new() };
            
            vybe_host::fetch_columns_for_query(&rc_str(da_conn), &final_query).unwrap_or_default()
        };

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
                    ControlType::RichTextBox => {
                        ui.label("HtmlText");
                        ui.vertical(|ui| {
                            ui.horizontal(|ui| {
                                if ui.small_button(egui::RichText::new("B").strong()).on_hover_text("Bold").clicked() { text.push_str("<b></b>"); }
                                if ui.small_button(egui::RichText::new("I").italics()).on_hover_text("Italic").clicked() { text.push_str("<i></i>"); }
                                if ui.small_button(egui::RichText::new("U").underline()).on_hover_text("Underline").clicked() { text.push_str("<u></u>"); }
                                if ui.small_button("A").on_hover_text("Color").clicked() { text.push_str("<font color=\"#ff0000\"></font>"); }
                            });
                            ui.add(egui::TextEdit::multiline(&mut text).desired_rows(4).desired_width(180.0));
                        });
                        ui.end_row();
                    }
                    ControlType::Label | ControlType::Button |
                    ControlType::CheckBox | ControlType::RadioButton |
                    ControlType::LinkLabel | ControlType::TextBox |
                    ControlType::MaskedTextBox |
                    ControlType::ListBox | ControlType::ComboBox |
                    ControlType::TabControl | ControlType::Frame |
                    ControlType::Panel => {
                        if ct == ControlType::TextBox && multiline {
                            ui.label("Text");
                            ui.add(egui::TextEdit::multiline(&mut text).desired_rows(3).desired_width(180.0));
                            ui.end_row();
                        } else {
                            row_str(ui, "Text",      &mut text);      ui.end_row();
                        }
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
                    | ControlType::ListBox | ControlType::ComboBox => {
                        // For Complex Binding Controls: Database connectivity
                        section_header(ui, "Data");
                        row_combo(ui, "DataSource", &mut data_source, &form_binding_sources); ui.end_row();
                        
                        let cols = resolve_columns_for_bs(state, &data_source, &resolve_conn_str);
                        row_combo(ui, "DataMember", &mut data_member, &cols); ui.end_row();
                        
                        if matches!(ct, ControlType::ComboBox | ControlType::ListBox) {
                            row_combo(ui, "DisplayMember", &mut display_member, &cols); ui.end_row();
                            row_combo(ui, "ValueMember",   &mut value_member, &cols);   ui.end_row();
                        }
                    }
                    _ => {}
                }

                // ── Advanced Data Bindings (Visual controls) ────────────
                if !matches!(ct, ControlType::DataGridView | ControlType::ListView | ControlType::ListBox | ControlType::ComboBox | ControlType::BindingNavigator) {
                    section_header(ui, "Data Bindings");
                    row_combo(ui, "DataSource", &mut db_source, &form_binding_sources); ui.end_row();
                    
                    let available_cols = resolve_columns_for_bs(state, &db_source, &resolve_conn_str);
                    
                    match ct {
                        ControlType::CheckBox | ControlType::RadioButton => {
                            row_combo(ui, "Checked", &mut db_checked, &available_cols); ui.end_row();
                            row_combo(ui, "Text",    &mut db_text, &available_cols);    ui.end_row();
                        }
                        ControlType::NumericUpDown | ControlType::TrackBar | ControlType::ProgressBar | ControlType::DateTimePicker => {
                            row_combo(ui, "Value",   &mut db_value, &available_cols);   ui.end_row();
                            row_combo(ui, "Text",    &mut db_text, &available_cols);    ui.end_row();
                        }
                        _ => {
                            row_combo(ui, "Text",    &mut db_text, &available_cols);    ui.end_row();
                        }
                    }
                }
            }

            // ── Non-visual component properties ─────────────────────────
            match ct {
                ControlType::BindingSourceComponent => {
                    section_header(ui, "Data Binding");
                    row_combo(ui, "DataSource", &mut data_source, &form_data_sources); ui.end_row();
                    
                    // If bound to a DataAdapter, let user pick the DataMember from available tables
                    let da_tables: Vec<String> = state.current_form_data()
                        .and_then(|f| f.controls.iter()
                            .find(|c| c.name.eq_ignore_ascii_case(&data_source) && matches!(c.control_type, ControlType::DataAdapterComponent)))
                        .and_then(|c| c.properties.get_string("ConnectionString").map(|s| resolve_conn_str(s)))
                        .and_then(|cs| vybe_host::test_connection_and_list_tables(&cs).ok())
                        .unwrap_or_default();
                    row_combo(ui, "DataMember", &mut data_member, &da_tables); ui.end_row();
                    
                    row_str(ui, "Filter",     &mut filter);      ui.end_row();
                    row_str(ui, "Sort",       &mut sort);        ui.end_row();
                }
                ControlType::DataAdapterComponent => {
                    section_header(ui, "Database");
                    row_str(ui, "SelectCommand",    &mut select_cmd);  ui.end_row();
                    row_str(ui, "ConnectionString", &mut conn_str);    ui.end_row();
                    
                    ui.label("Connection");
                    let mut lock = local_state_arc.lock().unwrap();
                    if ui.button(if lock.show_conn_builder { "▲ Hide Builder" } else { "🔧 Builder..." }).clicked() {
                        lock.show_conn_builder = !lock.show_conn_builder;
                    }
                    ui.end_row();
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

        // ── DataAdapter Connection Builder Sub-Panel ────────────────────
        if ct == ControlType::DataAdapterComponent {
            let mut lock = local_state_arc.lock().unwrap();
            let mut _cs_changed = false;
            let mut test_clicked = false;
            if lock.show_conn_builder {
                ui.group(|ui| {
                    ui.label(egui::RichText::new("🔧 Connection Builder").strong().color(egui::Color32::from_rgb(0, 120, 212)));
                    egui::Grid::new("conn_grid").num_columns(2).striped(true).show(ui, |ui| {
                        
                        ui.label("Server");
                        let mut type_changed = false;
                        egui::ComboBox::from_id_salt("dbtype_combo").selected_text(&db_type).show_ui(ui, |ui| {
                            if ui.selectable_label(db_type == "SQLite", "SQLite").clicked() { db_type = "SQLite".to_string(); type_changed = true; }
                            if ui.selectable_label(db_type == "PostgreSQL", "PostgreSQL").clicked() { db_type = "PostgreSQL".to_string(); type_changed = true; }
                            if ui.selectable_label(db_type == "MySQL", "MySQL").clicked() { db_type = "MySQL".to_string(); type_changed = true; }
                        });
                        ui.end_row();
                        
                        if type_changed {
                            if db_type == "PostgreSQL" { db_port = "5432".to_string(); }
                            else if db_type == "MySQL" { db_port = "3306".to_string(); }
                            else { db_port = "".to_string(); }
                            ui.ctx().request_repaint();
                        }
                        
                        if db_type == "SQLite" {
                            row_str(ui, "File", &mut db_path); ui.end_row();
                        } else {
                            row_str(ui, "Host", &mut db_host); ui.end_row();
                            row_str(ui, "Port", &mut db_port); ui.end_row();
                            row_str(ui, "Database", &mut db_name); ui.end_row();
                            row_str(ui, "User", &mut db_user); ui.end_row();
                            ui.label("Password"); ui.add(egui::TextEdit::singleline(&mut db_pass).password(true)); ui.end_row();
                        }
                    });
                    if ui.button("Build Connection String").clicked() {
                        let new_cs = match db_type.as_str() {
                            "SQLite" => format!("Data Source={}", if db_path.is_empty() { "database.db" } else { &db_path }),
                            "PostgreSQL" => format!("Host={};Port={};Database={};Username={};Password={}", db_host, db_port, db_name, db_user, db_pass),
                            "MySQL" => format!("Server={};Port={};Database={};Uid={};Pwd={}", db_host, db_port, db_name, db_user, db_pass),
                            _ => String::new(),
                        };
                        conn_str = new_cs;
                        _cs_changed = true;
                        lock.show_conn_builder = false;
                        ui.ctx().request_repaint();
                    }
                });
            }
            if ui.button("⚡ Test Connection & Fetch Tables").clicked() {
                test_clicked = true;
                ui.ctx().request_repaint();
            }
            if !lock.conn_status.is_empty() {
                ui.label(egui::RichText::new(&lock.conn_status).color(if lock.conn_status.starts_with('✓') { egui::Color32::from_rgb(0, 150, 0) } else { egui::Color32::from_rgb(200, 0, 0) }));
            }
            if !lock.da_tables.is_empty() {
                let mut sel_table = String::new();
                egui::ComboBox::from_id_salt("Query Builder").selected_text("— select a table —").show_ui(ui, |ui| {
                    for t in &lock.da_tables {
                        if ui.selectable_label(false, t).clicked() { sel_table = t.clone(); }
                    }
                });
                if !sel_table.is_empty() {
                    select_cmd = format!("SELECT * FROM {}", sel_table);
                    ui.ctx().request_repaint();
                }
            }
            
            // Drop locks before external call
            drop(lock);
            if test_clicked {
                let abs_cs = resolve_conn_str(&conn_str);
                let mut lock = local_state_arc.lock().unwrap();
                match vybe_host::test_connection_and_list_tables(&abs_cs) {
                    Ok(mut tbls) => {
                        tbls.sort();
                        lock.conn_status = format!("✓ Connected — {} tables found", tbls.len());
                        lock.da_tables = tbls;
                    }
                    Err(e) => {
                        lock.conn_status = format!("✗ {}", e);
                        lock.da_tables.clear();
                    }
                }
            }
        }

        // ── Write back all changes ───────────────────────────────────────
        if let Some(form) = state.current_form_data_mut() {
            if let Some(ctrl) = form.controls.iter_mut().find(|c| c.id == id) {
                ctrl.name = name;
                if !is_non_visual {
                    ctrl.properties.set("Text",      text);
                    ctrl.properties.set("BackColor",  back_color);
                    ctrl.properties.set("ForeColor",  fore_color);
                    ctrl.properties.set("Font",       font);
                    ctrl.properties.set("Enabled",    enabled);
                    ctrl.properties.set("Visible",    visible);
                    ctrl.properties.set("Checked",    checked);
                    ctrl.properties.set("ReadOnly",   read_only);
                    ctrl.properties.set("Multiline",  multiline);
                    ctrl.properties.set("BorderStyle", border_style);
                    ctrl.properties.set("Minimum",    minimum);
                    ctrl.properties.set("Maximum",    maximum);
                    ctrl.properties.set("Value",      value);
                    ctrl.properties.set("URL",        url);
                    ctrl.properties.set("LinkColor",  link_color);
                    
                    if matches!(ctrl.control_type, ControlType::DataGridView | ControlType::ListView | ControlType::ListBox | ControlType::ComboBox) {
                        ctrl.properties.set("DataSource",    data_source.clone());
                        ctrl.properties.set("DataMember",    data_member.clone());
                        ctrl.properties.set("DisplayMember", display_member);
                        ctrl.properties.set("ValueMember",   value_member);
                    } else {
                        ctrl.properties.set("DataBindings.Source",        db_source);
                        ctrl.properties.set("DataBindings.Text",          db_text);
                        ctrl.properties.set("DataBindings.Checked",       db_checked);
                        ctrl.properties.set("DataBindings.Value",         db_value);
                        ctrl.properties.set("DataBindings.SelectedValue", db_selected_value);
                        ctrl.properties.set("DataBindings.Visible",       db_visible);
                        ctrl.properties.set("DataBindings.Enabled",       db_enabled);
                    }
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
                        ctrl.properties.set("DbType",           db_type);
                        ctrl.properties.set("DbPath",           db_path);
                        ctrl.properties.set("DbHost",           db_host);
                        ctrl.properties.set("DbPort",           db_port);
                        ctrl.properties.set("DbName",           db_name);
                        ctrl.properties.set("DbUser",           db_user);
                        ctrl.properties.set("DbPassword",       db_pass);
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
                        ctrl.properties.set("Enabled",  enabled);
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

fn row_combo(ui: &mut Ui, label: &str, val: &mut String, options: &[String]) {
    ui.label(label);
    let current = if val.is_empty() { "(none)".to_string() } else { val.clone() };
    let mut changed = false;
    egui::ComboBox::from_id_salt(label)
        .selected_text(current)
        .show_ui(ui, |ui| {
            if ui.selectable_label(val.is_empty(), "(none)").clicked() { *val = String::new(); changed = true; }
            for opt in options {
                if ui.selectable_label(*val == *opt, opt).clicked() { *val = opt.clone(); changed = true; }
            }
        });
    if changed {
        ui.ctx().request_repaint();
    }
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
