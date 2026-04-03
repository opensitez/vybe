//! Properties panel — Properties tab + Events tab, matching legacy editor.
//! Reusable widget that lives in vybe_widgets so all IDE crates can share it.

use cosmic_text::{Color as CosmicColor, FontSystem, SwashCache};
use tiny_skia::{Paint, Pixmap, Transform};
use vybe_forms::{Form, ControlType};

use crate::layout::LayoutRect;
use crate::ide_text::{draw_text, measure_text};
use crate::color_picker::{ColorPicker, ColorPickerEvent, PickedColor};
use crate::font_picker::{FontPicker, FontPickerEvent};
use crate::dropdown::{Dropdown, DropdownEvent};

#[derive(Clone)]
pub struct EditingProp {
    pub key: String,
    pub value: String,
    pub cursor: usize,
}

#[derive(Clone, Copy, PartialEq)]
pub enum PropTab { Properties, Events }

const HEADER_H: f32 = 28.0;
const TAB_H: f32 = 26.0;
const ROW_H: f32 = 24.0;
const SECTION_H: f32 = 20.0;
const SCROLLBAR_W: f32 = 10.0;

/// A property row or section header.
enum PropItem {
    Section(String),
    Row(String, String),
    CheckboxRow(String, bool),
    DropdownRow(String, String, Vec<String>),
}

pub struct PropertiesPanel {
    pub scroll_y: f32,
    pub editing: Option<EditingProp>,
    pub tab: PropTab,
    /// Set when user clicks an event — (control_name, event_name).
    pub pending_event: Option<(String, String)>,
    /// Color picker popup state.
    pub color_picker: ColorPicker,
    /// Which property the color picker is editing ("BackColor" or "ForeColor").
    pub color_picker_prop: Option<String>,
    /// Font picker popup state.
    pub font_picker: FontPicker,
    /// Whether font picker is active.
    pub font_picker_active: bool,
    /// Connection wizard: whether the builder is expanded.
    pub conn_builder_open: bool,
    /// Connection wizard: test result message.
    pub conn_status: String,
    /// Connection wizard: table list from last test.
    pub conn_tables: Vec<String>,
    /// Pending action: "build_conn" or "test_conn" for the host to handle.
    pub pending_action: Option<String>,
    /// Set to true when an inline property edit should be committed immediately.
    pub pending_commit: bool,
    /// Dropdown popup state.
    pub dropdown: Option<Dropdown>,
    pub dropdown_prop: Option<String>,
    pub dropdown_pos: Option<(f32, f32)>,
    /// Properties being edited in the connection wizard modal.
    pub wizard_props: std::collections::HashMap<String, String>,
}

impl PropertiesPanel {
    pub fn new() -> Self {
        Self {
            scroll_y: 0.0,
            editing: None,
            tab: PropTab::Properties,
            pending_event: None,
            color_picker: ColorPicker::new(),
            color_picker_prop: None,
            font_picker: FontPicker::new(),
            font_picker_active: false,
            conn_builder_open: false,
            conn_status: String::new(),
            conn_tables: Vec::new(),
            pending_action: None,
            pending_commit: false,
            dropdown: None,
            dropdown_prop: None,
            dropdown_pos: None,
            wizard_props: std::collections::HashMap::new(),
        }
    }

    pub fn render(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        rect: LayoutRect, scale: f32, form: Option<&Form>, selected_control: Option<&str>,
    ) {
        let s = scale;
        let mut paint = Paint::default();

        // Background
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);

        // Left border
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y, 1.0, rect.h, s);

        // Title
        let title_color = CosmicColor::rgba(50, 50, 50, 255);
        draw_text(pix, fs, sc, "Properties", rect.x + 10.0, rect.y + 6.0, 13.0, title_color, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + HEADER_H - 1.0, rect.w, 1.0, s);

        // Tabs
        let tab_y = rect.y + HEADER_H;
        let tab_w = (rect.w - 2.0) / 2.0;
        let text_color = CosmicColor::rgba(30, 30, 30, 255);
        let dim_color = CosmicColor::rgba(100, 100, 100, 255);

        // Properties tab
        if self.tab == PropTab::Properties {
            paint.set_color_rgba8(227, 242, 253, 255);
        } else {
            paint.set_color_rgba8(245, 245, 245, 255);
        }
        fill(pix, &paint, rect.x + 1.0, tab_y, tab_w, TAB_H, s);
        draw_text(pix, fs, sc, "Properties", rect.x + 14.0, tab_y + 5.0, 12.0,
            if self.tab == PropTab::Properties { text_color } else { dim_color }, s);

        // Events tab
        if self.tab == PropTab::Events {
            paint.set_color_rgba8(227, 242, 253, 255);
        } else {
            paint.set_color_rgba8(245, 245, 245, 255);
        }
        fill(pix, &paint, rect.x + 1.0 + tab_w, tab_y, tab_w, TAB_H, s);
        draw_text(pix, fs, sc, "\u{26A1} Events", rect.x + tab_w + 14.0, tab_y + 5.0, 12.0,
            if self.tab == PropTab::Events { text_color } else { dim_color }, s);

        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, tab_y + TAB_H - 1.0, rect.w, 1.0, s);

        // Content area
        let content_top = tab_y + TAB_H;
        let content_h = rect.h - HEADER_H - TAB_H;
        let items = match self.tab {
            PropTab::Properties => self.collect_props(form, selected_control),
            PropTab::Events => self.collect_events(form, selected_control),
        };

        let label_color = CosmicColor::rgba(80, 80, 80, 255);
        let val_color = CosmicColor::rgba(30, 30, 30, 255);
        let section_color = CosmicColor::rgba(100, 100, 100, 255);
        let val_x = rect.x + rect.w * 0.42;
        let mut y = content_top + 2.0 - self.scroll_y;
        let mut entry_idx = 0usize;

        for item in &items {
            match item {
                PropItem::Section(label) => {
                    if y + SECTION_H > content_top && y < content_top + content_h {
                        paint.set_color_rgba8(240, 240, 240, 255);
                        fill(pix, &paint, rect.x + 1.0, y, rect.w - 2.0, SECTION_H, s);
                        draw_text(pix, fs, sc, label, rect.x + 8.0, y + 3.0, 10.0, section_color, s);
                    }
                    y += SECTION_H;
                }
                PropItem::Row(key, value) => {
                    if y + ROW_H > content_top && y < content_top + content_h {
                        if entry_idx % 2 == 1 {
                            paint.set_color_rgba8(247, 247, 247, 255);
                            fill(pix, &paint, rect.x + 1.0, y, rect.w - 2.0, ROW_H, s);
                        }

                        draw_text(pix, fs, sc, key, rect.x + 8.0, y + 4.0, 11.0, label_color, s);

                        paint.set_color_rgba8(230, 230, 230, 255);
                        fill(pix, &paint, val_x - 2.0, y, 1.0, ROW_H, s);

                        if let Some(ref ed) = self.editing {
                            if ed.key == *key {
                                paint.set_color_rgba8(255, 255, 255, 255);
                                fill(pix, &paint, val_x, y + 1.0, rect.w - (val_x - rect.x) - SCROLLBAR_W - 2.0, ROW_H - 2.0, s);
                                paint.set_color_rgba8(0, 120, 212, 255);
                                stroke_rect(pix, &paint, val_x, y + 1.0, rect.w - (val_x - rect.x) - SCROLLBAR_W - 2.0, ROW_H - 2.0, s);
                                draw_text(pix, fs, sc, &ed.value, val_x + 4.0, y + 4.0, 11.0, val_color, s);
                                let text_up_to_cursor = &ed.value[0..ed.cursor];
                                let w = measure_text(fs, text_up_to_cursor, 11.0, s);
                                let cx = val_x + 4.0 + w;
                                paint.set_color_rgba8(0, 0, 0, 255);
                                fill(pix, &paint, cx, y + 3.0, 1.0, ROW_H - 6.0, s);
                            } else {
                                draw_text(pix, fs, sc, value, val_x + 4.0, y + 4.0, 11.0, val_color, s);
                            }
                        } else {
                            draw_text(pix, fs, sc, value, val_x + 4.0, y + 4.0, 11.0, val_color, s);
                        }

                        // Color swatch for BackColor / ForeColor
                        if key == "BackColor" || key == "ForeColor" {
                            let swatch_x = rect.x + rect.w - SCROLLBAR_W - 22.0;
                            let swatch_y = y + 3.0;
                            let swatch_sz = ROW_H - 6.0;
                            if let Some(c) = PickedColor::from_hex(value) {
                                let mut sp = Paint::default();
                                sp.set_color_rgba8(c.r, c.g, c.b, c.a);
                                fill(pix, &sp, swatch_x, swatch_y, swatch_sz, swatch_sz, s);
                                sp.set_color_rgba8(160, 160, 160, 255);
                                stroke_rect(pix, &sp, swatch_x, swatch_y, swatch_sz, swatch_sz, s);
                            }
                        }

                        paint.set_color_rgba8(235, 235, 235, 255);
                        fill(pix, &paint, rect.x + 1.0, y + ROW_H - 1.0, rect.w - 2.0, 1.0, s);
                    }
                    entry_idx += 1;
                    y += ROW_H;
                }
                PropItem::CheckboxRow(key, checked) => {
                    if y + ROW_H > content_top && y < content_top + content_h {
                        if entry_idx % 2 == 1 {
                            paint.set_color_rgba8(247, 247, 247, 255);
                            fill(pix, &paint, rect.x + 1.0, y, rect.w - 2.0, ROW_H, s);
                        }
                        draw_text(pix, fs, sc, key, rect.x + 8.0, y + 4.0, 11.0, label_color, s);
                        paint.set_color_rgba8(230, 230, 230, 255);
                        fill(pix, &paint, val_x - 2.0, y, 1.0, ROW_H, s);
                        let cb_x = val_x + 4.0;
                        let cb_y = y + 4.0;
                        let cb_sz = 14.0;
                        paint.set_color_rgba8(255, 255, 255, 255);
                        fill(pix, &paint, cb_x, cb_y, cb_sz, cb_sz, s);
                        paint.set_color_rgba8(160, 160, 160, 255);
                        stroke_rect(pix, &paint, cb_x, cb_y, cb_sz, cb_sz, s);
                        if *checked {
                            paint.set_color_rgba8(0, 120, 212, 255);
                            fill(pix, &paint, cb_x + 3.0, cb_y + 3.0, cb_sz - 6.0, cb_sz - 6.0, s);
                        }
                        let label_text = if *checked { "True" } else { "False" };
                        draw_text(pix, fs, sc, label_text, cb_x + cb_sz + 4.0, y + 4.0, 11.0, val_color, s);
                        paint.set_color_rgba8(235, 235, 235, 255);
                        fill(pix, &paint, rect.x + 1.0, y + ROW_H - 1.0, rect.w - 2.0, 1.0, s);
                    }
                    entry_idx += 1;
                    y += ROW_H;
                }
                PropItem::DropdownRow(key, current, _options) => {
                    if y + ROW_H > content_top && y < content_top + content_h {
                        if entry_idx % 2 == 1 {
                            paint.set_color_rgba8(247, 247, 247, 255);
                            fill(pix, &paint, rect.x + 1.0, y, rect.w - 2.0, ROW_H, s);
                        }
                        draw_text(pix, fs, sc, key, rect.x + 8.0, y + 4.0, 11.0, label_color, s);
                        paint.set_color_rgba8(230, 230, 230, 255);
                        fill(pix, &paint, val_x - 2.0, y, 1.0, ROW_H, s);
                        let dd_w = rect.w - (val_x - rect.x) - SCROLLBAR_W - 2.0;
                        paint.set_color_rgba8(255, 255, 255, 255);
                        fill(pix, &paint, val_x, y + 1.0, dd_w, ROW_H - 2.0, s);
                        paint.set_color_rgba8(180, 180, 180, 255);
                        stroke_rect(pix, &paint, val_x, y + 1.0, dd_w, ROW_H - 2.0, s);
                        draw_text(pix, fs, sc, current, val_x + 4.0, y + 4.0, 11.0, val_color, s);
                        let arrow_x = val_x + dd_w - 14.0;
                        let arrow_y = y + ROW_H / 2.0 - 2.0;
                        paint.set_color_rgba8(80, 80, 80, 255);
                        fill(pix, &paint, arrow_x, arrow_y, 8.0, 1.0, s);
                        fill(pix, &paint, arrow_x + 1.0, arrow_y + 1.0, 6.0, 1.0, s);
                        fill(pix, &paint, arrow_x + 2.0, arrow_y + 2.0, 4.0, 1.0, s);
                        fill(pix, &paint, arrow_x + 3.0, arrow_y + 3.0, 2.0, 1.0, s);
                        paint.set_color_rgba8(235, 235, 235, 255);
                        fill(pix, &paint, rect.x + 1.0, y + ROW_H - 1.0, rect.w - 2.0, 1.0, s);
                    }
                    entry_idx += 1;
                    y += ROW_H;
                }
            }
        }

        if items.is_empty() {
            draw_text(pix, fs, sc, "No selection", rect.x + 10.0, content_top + 8.0, 12.0, dim_color, s);
        }

        // Overdraw content top to clip scrolled items
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, HEADER_H + TAB_H, s);
        // Re-render header+tabs
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, HEADER_H, s);
        draw_text(pix, fs, sc, "Properties", rect.x + 10.0, rect.y + 6.0, 13.0, title_color, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + HEADER_H - 1.0, rect.w, 1.0, s);
        if self.tab == PropTab::Properties {
            paint.set_color_rgba8(227, 242, 253, 255);
        } else {
            paint.set_color_rgba8(245, 245, 245, 255);
        }
        fill(pix, &paint, rect.x + 1.0, tab_y, tab_w, TAB_H, s);
        draw_text(pix, fs, sc, "Properties", rect.x + 14.0, tab_y + 5.0, 12.0,
            if self.tab == PropTab::Properties { text_color } else { dim_color }, s);
        if self.tab == PropTab::Events {
            paint.set_color_rgba8(227, 242, 253, 255);
        } else {
            paint.set_color_rgba8(245, 245, 245, 255);
        }
        fill(pix, &paint, rect.x + 1.0 + tab_w, tab_y, tab_w, TAB_H, s);
        draw_text(pix, fs, sc, "\u{26A1} Events", rect.x + tab_w + 14.0, tab_y + 5.0, 12.0,
            if self.tab == PropTab::Events { text_color } else { dim_color }, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, tab_y + TAB_H - 1.0, rect.w, 1.0, s);

        // Scrollbar
        let total_h = self.items_height(&items);
        let max_scroll = (total_h - content_h).max(0.0);
        if max_scroll > 0.0 {
            let sb_x = rect.x + rect.w - SCROLLBAR_W;
            paint.set_color_rgba8(235, 235, 235, 255);
            fill(pix, &paint, sb_x, content_top, SCROLLBAR_W, content_h, s);
            let visible_frac = (content_h / total_h).min(1.0);
            let thumb_h = (content_h * visible_frac).max(20.0);
            let scroll_frac = if max_scroll > 0.0 { self.scroll_y / max_scroll } else { 0.0 };
            let thumb_y = content_top + scroll_frac * (content_h - thumb_h);
            paint.set_color_rgba8(190, 190, 190, 255);
            fill(pix, &paint, sb_x + 2.0, thumb_y, SCROLLBAR_W - 4.0, thumb_h, s);
        }

        // Left border (re-draw on top)
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y, 1.0, rect.h, s);

        // Color picker popup overlay
        if self.color_picker.open {
            let popup_x = rect.x + 10.0;
            let popup_y = rect.y + HEADER_H + TAB_H + 40.0;
            self.color_picker.render_popup(pix, popup_x, popup_y, s);
        }

        // Font picker popup overlay
        if self.font_picker.open {
            let popup_x = rect.x + 10.0;
            let popup_y = rect.y + HEADER_H + TAB_H + 40.0;
            self.font_picker.render_popup(pix, fs, sc, popup_x, popup_y, s);
        }

        // Dropdown popup overlay
        if let Some(ref dropdown) = self.dropdown {
            let (px, py) = self.dropdown_pos.unwrap_or((rect.x + 10.0, rect.y + HEADER_H + TAB_H + 40.0));
            dropdown.render_list(
                pix, fs, sc, px, py,
                (252, 252, 252, 255), (180, 180, 180, 255),
                (0, 120, 212, 40), (0, 120, 212, 25),
                CosmicColor::rgba(0, 90, 180, 255),
                CosmicColor::rgba(30, 30, 30, 255),
            );
        }

        // Connection Wizard Overlay
        self.render_conn_wizard(pix, fs, sc, rect, s);
    }

    fn render_conn_wizard(&self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, rect: LayoutRect, s: f32) {
        if !self.conn_builder_open { return; }
        let mut paint = Paint::default();
        paint.set_color_rgba8(0, 0, 0, 120);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);

        let pw = rect.w - 20.0;
        let _ph = 210.0;
        let px = rect.x + 10.0;
        let py = rect.y + HEADER_H + 30.0;

        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, px, py, pw, _ph, s);
        paint.set_color_rgba8(160, 160, 160, 255);
        stroke_rect(pix, &paint, px, py, pw, _ph, s);

        let text_col = CosmicColor::rgba(30, 30, 30, 255);
        let header_col = CosmicColor::rgba(0, 120, 212, 255);
        draw_text(pix, fs, sc, "Connection Builder", px + 10.0, py + 10.0, 13.0, header_col, s);

        // Close button (X)
        draw_text(pix, fs, sc, "\u{2715}", px + pw - 20.0, py + 10.0, 14.0, CosmicColor::rgba(180, 50, 50, 255), s);

        let db_type = self.wizard_props.get("DbType").map(|s| s.as_str()).unwrap_or("PostgreSQL");
        let fields = if db_type == "SQLite" {
            vec![("DbPath", "Database File")]
        } else {
            vec![
                ("DbHost", "Host"),
                ("DbPort", "Port"),
                ("DbName", "Database"),
                ("DbUser", "Username"),
                ("DbPassword", "Password"),
            ]
        };

        let mut cy = py + 35.0;
        for (key, label) in fields.iter() {
            draw_text(pix, fs, sc, label, px + 10.0, cy + 4.0, 11.0, text_col, s);

            let ix = px + 70.0;
            let iw = pw - 80.0;
            paint.set_color_rgba8(255, 255, 255, 255);
            fill(pix, &paint, ix, cy, iw, 20.0, s);
            paint.set_color_rgba8(160, 160, 160, 255);
            stroke_rect(pix, &paint, ix, cy, iw, 20.0, s);

            let val = self.wizard_props.get(*key).map(|st| st.as_str()).unwrap_or("");
            let display = if let Some(ed) = &self.editing {
                if ed.key == format!("Wizard_{}", key) { ed.value.as_str() } else { val }
            } else { val };

            let display_pw = if *key == "DbPassword" { "\u{2022}".repeat(display.chars().count()) } else { display.to_string() };
            draw_text(pix, fs, sc, &display_pw, ix + 4.0, cy + 3.0, 11.0, text_col, s);

            if let Some(ed) = &self.editing {
                if ed.key == format!("Wizard_{}", key) {
                    let chars_before = ed.value[0..ed.cursor].chars().count();
                    let text_up_to_cursor = if *key == "DbPassword" {
                        "\u{2022}".repeat(chars_before)
                    } else {
                        ed.value[0..ed.cursor].to_string()
                    };
                    let w = measure_text(fs, &text_up_to_cursor, 11.0, s);
                    paint.set_color_rgba8(0, 0, 0, 255);
                    fill(pix, &paint, ix + 4.0 + w, cy + 2.0, 1.0, 16.0, s);
                }
            }
            cy += 24.0;
        }

        // Build Button
        let bx = px + 10.0;
        let by = cy + 10.0;
        let bw = pw - 20.0;
        paint.set_color_rgba8(0, 120, 212, 255);
        fill(pix, &paint, bx, by, bw, 24.0, s);
        draw_text(pix, fs, sc, "Build Connection String", bx + bw / 2.0 - 70.0, by + 5.0, 11.0, CosmicColor::rgba(255, 255, 255, 255), s);
    }

    fn items_height(&self, items: &[PropItem]) -> f32 {
        items.iter().map(|i| match i {
            PropItem::Section(_) => SECTION_H,
            PropItem::Row(_, _) | PropItem::CheckboxRow(_, _) | PropItem::DropdownRow(_, _, _) => ROW_H,
        }).sum()
    }

    fn prop(key: &str, val: &str) -> PropItem {
        PropItem::Row(key.into(), val.into())
    }

    fn get_prop(ctrl: &vybe_forms::Control, key: &str, default: &str) -> String {
        ctrl.properties.get_string(key).unwrap_or(default).to_string()
    }

    fn collect_props(&self, form: Option<&Form>, selected_control: Option<&str>) -> Vec<PropItem> {
        use ControlType::*;
        if let (Some(form), Some(ctrl_name)) = (form, selected_control) {
            if let Some(ctrl) = form.controls.iter().find(|c| c.name == ctrl_name) {
                let mut items = vec![
                    PropItem::Section("Basic".into()),
                    Self::prop("Name", &ctrl.name),
                    Self::prop("Type", &format!("{:?}", ctrl.control_type)),
                ];

                if !ctrl.control_type.is_non_visual() {
                    items.push(Self::prop("Left", &format!("{}", ctrl.bounds.x)));
                    items.push(Self::prop("Top", &format!("{}", ctrl.bounds.y)));
                    items.push(Self::prop("Width", &format!("{}", ctrl.bounds.width)));
                    items.push(Self::prop("Height", &format!("{}", ctrl.bounds.height)));
                }

                // Appearance
                items.push(PropItem::Section("Appearance".into()));
                items.push(Self::prop("Text", &Self::get_prop(ctrl, "Text", "")));
                items.push(Self::prop("BackColor", &Self::get_prop(ctrl, "BackColor", "")));
                items.push(Self::prop("ForeColor", &Self::get_prop(ctrl, "ForeColor", "")));
                items.push(Self::prop("Font", &Self::get_prop(ctrl, "Font", "")));
                items.push(Self::prop("TabIndex", &format!("{}", ctrl.tab_index)));
                items.push(PropItem::CheckboxRow("Enabled".into(),
                    Self::get_prop(ctrl, "Enabled", "True").eq_ignore_ascii_case("true")));
                items.push(PropItem::CheckboxRow("Visible".into(),
                    Self::get_prop(ctrl, "Visible", "True").eq_ignore_ascii_case("true")));

                // Control-specific
                match ctrl.control_type {
                    CheckBox | RadioButton => {
                        items.push(PropItem::Section("State".into()));
                        items.push(PropItem::CheckboxRow("Checked".into(),
                            Self::get_prop(ctrl, "Checked", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::DropdownRow("CheckState".into(),
                            Self::get_prop(ctrl, "CheckState", "Unchecked"),
                            vec!["Unchecked".into(), "Checked".into(), "Indeterminate".into()]));
                    }
                    TextBox => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::CheckboxRow("Multiline".into(),
                            Self::get_prop(ctrl, "Multiline", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("ReadOnly".into(),
                            Self::get_prop(ctrl, "ReadOnly", "False").eq_ignore_ascii_case("true")));
                        items.push(Self::prop("MaxLength", &Self::get_prop(ctrl, "MaxLength", "0")));
                        items.push(Self::prop("PasswordChar", &Self::get_prop(ctrl, "PasswordChar", "")));
                    }
                    RichTextBox => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::CheckboxRow("ReadOnly".into(),
                            Self::get_prop(ctrl, "ReadOnly", "False").eq_ignore_ascii_case("true")));
                    }
                    NumericUpDown => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(Self::prop("Minimum", &Self::get_prop(ctrl, "Minimum", "0")));
                        items.push(Self::prop("Maximum", &Self::get_prop(ctrl, "Maximum", "100")));
                        items.push(Self::prop("Increment", &Self::get_prop(ctrl, "Increment", "1")));
                        items.push(Self::prop("DecimalPlaces", &Self::get_prop(ctrl, "DecimalPlaces", "0")));
                    }
                    TrackBar => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(Self::prop("Minimum", &Self::get_prop(ctrl, "Minimum", "0")));
                        items.push(Self::prop("Maximum", &Self::get_prop(ctrl, "Maximum", "10")));
                        items.push(Self::prop("SmallChange", &Self::get_prop(ctrl, "SmallChange", "1")));
                        items.push(Self::prop("LargeChange", &Self::get_prop(ctrl, "LargeChange", "5")));
                    }
                    ProgressBar => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(Self::prop("Minimum", &Self::get_prop(ctrl, "Minimum", "0")));
                        items.push(Self::prop("Maximum", &Self::get_prop(ctrl, "Maximum", "100")));
                    }
                    MaskedTextBox => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(Self::prop("Mask", &Self::get_prop(ctrl, "Mask", "")));
                        items.push(Self::prop("PromptChar", &Self::get_prop(ctrl, "PromptChar", "_")));
                    }
                    ComboBox | ListBox => {
                        items.push(PropItem::Section("Behavior".into()));
                        if ctrl.control_type == ComboBox {
                            items.push(PropItem::DropdownRow("DropDownStyle".into(),
                                Self::get_prop(ctrl, "DropDownStyle", "DropDown"),
                                vec!["Simple".into(), "DropDown".into(), "DropDownList".into()]));
                        }
                        if ctrl.control_type == ListBox {
                            items.push(PropItem::CheckboxRow("Sorted".into(),
                                Self::get_prop(ctrl, "Sorted", "False").eq_ignore_ascii_case("true")));
                            items.push(PropItem::DropdownRow("SelectionMode".into(),
                                Self::get_prop(ctrl, "SelectionMode", "One"),
                                vec!["None".into(), "One".into(), "MultiSimple".into(), "MultiExtended".into()]));
                        }
                    }
                    TreeView => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::CheckboxRow("CheckBoxes".into(),
                            Self::get_prop(ctrl, "CheckBoxes", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("ShowLines".into(),
                            Self::get_prop(ctrl, "ShowLines", "True").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("ShowRootLines".into(),
                            Self::get_prop(ctrl, "ShowRootLines", "True").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("ShowPlusMinus".into(),
                            Self::get_prop(ctrl, "ShowPlusMinus", "True").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("LabelEdit".into(),
                            Self::get_prop(ctrl, "LabelEdit", "False").eq_ignore_ascii_case("true")));
                    }
                    ListView => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::DropdownRow("View".into(),
                            Self::get_prop(ctrl, "View", "Details"),
                            vec!["LargeIcon".into(), "Details".into(), "SmallIcon".into(), "List".into(), "Tile".into()]));
                        items.push(PropItem::CheckboxRow("FullRowSelect".into(),
                            Self::get_prop(ctrl, "FullRowSelect", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("GridLines".into(),
                            Self::get_prop(ctrl, "GridLines", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("CheckBoxes".into(),
                            Self::get_prop(ctrl, "CheckBoxes", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("MultiSelect".into(),
                            Self::get_prop(ctrl, "MultiSelect", "True").eq_ignore_ascii_case("true")));
                    }
                    DataGridView => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::CheckboxRow("ReadOnly".into(),
                            Self::get_prop(ctrl, "ReadOnly", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("AllowUserToAddRows".into(),
                            Self::get_prop(ctrl, "AllowUserToAddRows", "True").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("AllowUserToDeleteRows".into(),
                            Self::get_prop(ctrl, "AllowUserToDeleteRows", "True").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("AutoGenerateColumns".into(),
                            Self::get_prop(ctrl, "AutoGenerateColumns", "True").eq_ignore_ascii_case("true")));
                    }
                    TabControl => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::DropdownRow("Alignment".into(),
                            Self::get_prop(ctrl, "Alignment", "Top"),
                            vec!["Top".into(), "Bottom".into(), "Left".into(), "Right".into()]));
                        items.push(PropItem::CheckboxRow("Multiline".into(),
                            Self::get_prop(ctrl, "Multiline", "False").eq_ignore_ascii_case("true")));
                    }
                    DateTimePicker => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::DropdownRow("Format".into(),
                            Self::get_prop(ctrl, "Format", "Long"),
                            vec!["Long".into(), "Short".into(), "Time".into(), "Custom".into()]));
                    }
                    WebBrowser => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(Self::prop("URL", &Self::get_prop(ctrl, "URL", "")));
                    }
                    _ => {}
                }

                // Data Binding
                items.push(PropItem::Section("Data".into()));

                let is_non_visual = ctrl.control_type.is_non_visual();
                let has_complex = matches!(ctrl.control_type,
                    DataGridView | ListBox | ComboBox | BindingNavigator) ||
                    matches!(ctrl.control_type, vybe_forms::ControlType::BindingSourceComponent);

                let mut bs_options = vec!["(none)".to_string()];
                let mut ds_options = vec!["(none)".to_string()];
                for c in &form.controls {
                    if matches!(c.control_type, vybe_forms::ControlType::BindingSourceComponent) && c.id != ctrl.id {
                        bs_options.push(c.name.clone());
                        ds_options.push(c.name.clone());
                    }
                    if matches!(c.control_type, vybe_forms::ControlType::DataAdapterComponent |
                        vybe_forms::ControlType::DataSetComponent |
                        vybe_forms::ControlType::DataTableComponent |
                        vybe_forms::ControlType::DataView) {
                        ds_options.push(c.name.clone());
                    }
                }

                if !is_non_visual && !has_complex {
                    let bindable = match ctrl.control_type {
                        CheckBox | RadioButton => "Checked",
                        PictureBox => "ImageLocation",
                        _ => "Text",
                    };
                    items.push(PropItem::DropdownRow("DataBindings.Source".into(),
                        Self::get_prop(ctrl, "DataBindings.Source", ""), bs_options.clone()));
                    items.push(Self::prop(&format!("Bind: {}", bindable),
                        &Self::get_prop(ctrl, &format!("DataBindings.{}", bindable), "")));
                }

                if has_complex && ctrl.control_type != BindingNavigator {
                    items.push(PropItem::DropdownRow("DataSource".into(),
                        Self::get_prop(ctrl, "DataSource", ""), ds_options.clone()));
                    items.push(Self::prop("DataMember", &Self::get_prop(ctrl, "DataMember", "")));
                }

                if matches!(ctrl.control_type, ListBox | ComboBox) {
                    items.push(Self::prop("DisplayMember", &Self::get_prop(ctrl, "DisplayMember", "")));
                    items.push(Self::prop("ValueMember", &Self::get_prop(ctrl, "ValueMember", "")));
                }

                if ctrl.control_type == BindingNavigator {
                    items.push(PropItem::DropdownRow("BindingSource".into(),
                        Self::get_prop(ctrl, "BindingSource", ""), bs_options.clone()));
                }

                if ctrl.control_type == BindingSourceComponent {
                    items.push(Self::prop("DataSource", &Self::get_prop(ctrl, "DataSource", "")));
                    items.push(Self::prop("DataMember", &Self::get_prop(ctrl, "DataMember", "")));
                    items.push(Self::prop("Filter", &Self::get_prop(ctrl, "Filter", "")));
                    items.push(Self::prop("Sort", &Self::get_prop(ctrl, "Sort", "")));
                }

                if ctrl.control_type == DataSetComponent {
                    items.push(Self::prop("DataSetName", &Self::get_prop(ctrl, "DataSetName", "NewDataSet")));
                }

                if ctrl.control_type == DataTableComponent {
                    items.push(Self::prop("TableName", &Self::get_prop(ctrl, "TableName", "Table1")));
                }

                if ctrl.control_type == DataAdapterComponent {
                    items.push(Self::prop("SelectCommand", &Self::get_prop(ctrl, "SelectCommand", "")));
                    items.push(Self::prop("ConnectionString", &Self::get_prop(ctrl, "ConnectionString", "")));
                    items.push(PropItem::Section("Connection Builder".into()));
                    items.push(PropItem::DropdownRow("DbType".into(),
                        Self::get_prop(ctrl, "DbType", "SQLite"),
                        vec!["SQLite".into(), "PostgreSQL".into(), "MySQL".into()]));
                    let db_type = Self::get_prop(ctrl, "DbType", "SQLite");
                    if db_type == "SQLite" {
                        items.push(Self::prop("DbPath", &Self::get_prop(ctrl, "DbPath", "")));
                    }
                    items.push(Self::prop("\u{26A1} Build ConnStr", "\u{2192} click"));
                    items.push(Self::prop("\u{1F50C} Test Connection", "\u{2192} click"));
                    if !self.conn_status.is_empty() {
                        items.push(Self::prop("Status", &self.conn_status));
                    }
                    for tbl in &self.conn_tables {
                        items.push(Self::prop("  Table", tbl));
                    }
                }

                return items;
            }
        } else if let Some(form) = form {
            return vec![
                PropItem::Section("Form".into()),
                Self::prop("Name", &form.name),
                Self::prop("Text", &form.text),
                Self::prop("Width", &format!("{}", form.width)),
                Self::prop("Height", &format!("{}", form.height)),
                PropItem::Section("Appearance".into()),
                Self::prop("BackColor", &form.back_color.clone().unwrap_or_default()),
                Self::prop("ForeColor", &form.fore_color.clone().unwrap_or_default()),
                Self::prop("Font", &form.font.clone().unwrap_or_default()),
            ];
        }
        vec![]
    }

    fn collect_events(&self, form: Option<&Form>, selected_control: Option<&str>) -> Vec<PropItem> {
        use ControlType::*;
        let ct = if let (Some(form), Some(name)) = (form, selected_control) {
            form.controls.iter().find(|c| c.name == name).map(|c| c.control_type.clone())
        } else if form.is_some() {
            None
        } else {
            return vec![];
        };

        let events: &[&str] = match ct.as_ref() {
            Some(Button) => &["Click", "MouseDown", "MouseUp", "MouseMove", "MouseEnter", "MouseLeave", "GotFocus", "LostFocus", "KeyDown", "KeyUp", "KeyPress", "EnabledChanged", "VisibleChanged", "Paint"],
            Some(TextBox) | Some(MaskedTextBox) | Some(RichTextBox) => &["TextChanged", "KeyPress", "KeyDown", "KeyUp", "GotFocus", "LostFocus", "Click", "Enter", "Leave", "Validating", "Validated"],
            Some(Label) | Some(LinkLabel) => &["Click", "DoubleClick", "MouseEnter", "MouseLeave"],
            Some(CheckBox) | Some(RadioButton) => &["CheckedChanged", "Click", "GotFocus", "LostFocus", "KeyPress", "EnabledChanged"],
            Some(ListBox) => &["SelectedIndexChanged", "Click", "DoubleClick", "GotFocus", "LostFocus", "KeyPress", "KeyDown"],
            Some(ComboBox) => &["SelectedIndexChanged", "SelectedValueChanged", "TextChanged", "DropDown", "DropDownClosed", "Click", "GotFocus", "LostFocus", "KeyPress", "KeyDown"],
            Some(TreeView) => &["AfterSelect", "BeforeSelect", "AfterExpand", "AfterCollapse", "NodeMouseClick", "NodeMouseDoubleClick", "AfterCheck", "Click", "DoubleClick", "KeyDown"],
            Some(ListView) => &["SelectedIndexChanged", "ItemActivate", "ColumnClick", "ItemCheck", "Click", "DoubleClick", "KeyDown"],
            Some(DataGridView) => &["CellClick", "CellDoubleClick", "CellValueChanged", "CellEndEdit", "CellBeginEdit", "SelectionChanged", "RowEnter", "RowLeave", "DataBindingComplete", "KeyDown"],
            Some(TabControl) => &["SelectedIndexChanged", "Selected", "Deselecting", "Click", "DoubleClick"],
            Some(Panel) | Some(Frame) => &["Click", "DoubleClick", "MouseDown", "MouseUp", "MouseMove", "MouseEnter", "MouseLeave", "Paint", "Resize"],
            Some(PictureBox) => &["Click", "DoubleClick", "MouseDown", "MouseUp", "MouseMove", "Paint"],
            Some(ProgressBar) => &["Click", "ValueChanged"],
            Some(NumericUpDown) => &["ValueChanged", "KeyPress", "KeyDown", "GotFocus", "LostFocus", "Validating"],
            Some(DateTimePicker) => &["ValueChanged", "DropDown", "DropDownClosed", "GotFocus", "LostFocus", "KeyPress"],
            Some(TrackBar) => &["Scroll", "ValueChanged", "MouseDown", "MouseUp", "GotFocus", "LostFocus"],
            Some(MenuStrip) | Some(ContextMenuStrip) => &["ItemClicked", "Click"],
            Some(StatusStrip) => &["ItemClicked", "Click"],
            Some(ToolStrip) => &["ItemClicked", "ButtonClick", "Click"],
            Some(WebBrowser) => &["DocumentCompleted", "Navigating", "Navigated", "ProgressChanged", "Click"],
            Some(SplitContainer) => &["SplitterMoved", "SplitterMoving", "Click", "DoubleClick"],
            Some(MonthCalendar) => &["DateChanged", "DateSelected", "Click", "DoubleClick", "GotFocus", "LostFocus"],
            Some(BindingSourceComponent) => &["CurrentChanged", "PositionChanged", "DataSourceChanged"],
            Some(BindingNavigator) => &["Click"],
            Some(Timer) => &["Tick"],
            None => &["Load", "Shown", "Activated", "Deactivate", "FormClosing", "FormClosed", "Resize", "Paint", "Click", "DoubleClick", "KeyDown", "KeyUp", "KeyPress", "MouseClick", "MouseDown", "MouseUp"],
            _ => &["Click", "DoubleClick", "MouseDown", "MouseUp", "GotFocus", "LostFocus"],
        };

        let mut items = Vec::new();
        for &ev in events {
            items.push(PropItem::Row(ev.into(), String::new()));
        }
        items
    }

    pub fn handle_click(&mut self, mx: f32, my: f32, rect: LayoutRect, form: Option<&Form>, selected_control: Option<&str>, scale: f32) {
        if !rect.contains(mx, my) { return; }

        if self.conn_builder_open {
            let pw = rect.w - 20.0;
            let _ph = 210.0;
            let px = rect.x + 10.0;
            let py = rect.y + HEADER_H + 30.0;

            // Close X
            if mx >= px + pw - 24.0 && mx <= px + pw && my >= py + 6.0 && my <= py + 26.0 {
                self.conn_builder_open = false;
                self.editing = None;
                return;
            }

            let bx = px + 10.0;
            let by = py + 35.0 + (5.0 * 24.0) + 10.0;
            let bw = pw - 20.0;
            // Build Button Check
            if mx >= bx && mx <= bx + bw && my >= by && my <= by + 24.0 {
                if let Some(ed) = self.editing.take() {
                    if ed.key.starts_with("Wizard_") {
                        let k = ed.key.replace("Wizard_", "");
                        self.wizard_props.insert(k, ed.value);
                    }
                }

                let db_type = self.wizard_props.get("DbType").map(String::as_str).unwrap_or("PostgreSQL");
                let conn_str = if db_type == "SQLite" {
                    let path = self.wizard_props.get("DbPath").map(String::as_str).unwrap_or("database.db");
                    format!("Data Source={}", path)
                } else if db_type == "MySQL" {
                    let host = self.wizard_props.get("DbHost").map(String::as_str).unwrap_or("localhost");
                    let port = self.wizard_props.get("DbPort").map(String::as_str).unwrap_or("3306");
                    let db = self.wizard_props.get("DbName").map(String::as_str).unwrap_or("");
                    let user = self.wizard_props.get("DbUser").map(String::as_str).unwrap_or("root");
                    let pass = self.wizard_props.get("DbPassword").map(String::as_str).unwrap_or("");
                    format!("Server={};Port={};Database={};Uid={};Pwd={}", host, port, db, user, pass)
                } else {
                    let host = self.wizard_props.get("DbHost").map(String::as_str).unwrap_or("localhost");
                    let port = self.wizard_props.get("DbPort").map(String::as_str).unwrap_or("5432");
                    let db = self.wizard_props.get("DbName").map(String::as_str).unwrap_or("");
                    let user = self.wizard_props.get("DbUser").map(String::as_str).unwrap_or("postgres");
                    let pass = self.wizard_props.get("DbPassword").map(String::as_str).unwrap_or("");
                    format!("Host={};Port={};Database={};Username={};Password={}", host, port, db, user, pass)
                };

                self.editing = Some(EditingProp {
                    key: "ConnectionString".into(),
                    value: conn_str,
                    cursor: 0,
                });
                self.pending_commit = true;
                self.conn_builder_open = false;
                return;
            }

            let db_type = self.wizard_props.get("DbType").map(String::as_str).unwrap_or("PostgreSQL");
            let fields = if db_type == "SQLite" {
                vec!["DbPath"]
            } else {
                vec!["DbHost", "DbPort", "DbName", "DbUser", "DbPassword"]
            };
            let mut cy = py + 35.0;
            for key in &fields {
                let ix = px + 70.0;
                let iw = pw - 80.0;
                if mx >= ix && mx <= ix + iw && my >= cy && my <= cy + 20.0 {
                    if let Some(ed) = self.editing.take() {
                        if ed.key.starts_with("Wizard_") {
                            let k = ed.key.replace("Wizard_", "");
                            self.wizard_props.insert(k, ed.value);
                        }
                    }
                    let val = self.wizard_props.get(*key).map(String::as_str).unwrap_or("").to_string();
                    self.editing = Some(EditingProp {
                        key: format!("Wizard_{}", key),
                        value: val.clone(),
                        cursor: val.len(),
                    });
                    return;
                }
                cy += 24.0;
            }

            if let Some(ed) = self.editing.take() {
                if ed.key.starts_with("Wizard_") {
                    let k = ed.key.replace("Wizard_", "");
                    self.wizard_props.insert(k, ed.value);
                }
            }
            return;
        }

        // Color picker popup is open: route click to it first
        if self.color_picker.open {
            let popup_x = rect.x + 10.0;
            let popup_y = rect.y + HEADER_H + TAB_H + 40.0;
            match self.color_picker.handle_click(mx, my, popup_x, popup_y) {
                ColorPickerEvent::Changed(c) => {
                    if let Some(ref prop_name) = self.color_picker_prop {
                        self.editing = Some(EditingProp {
                            key: prop_name.clone(),
                            value: c.to_hex(),
                            cursor: 7,
                        });
                    }
                    return;
                }
                ColorPickerEvent::Closed => {
                    if let Some(ref prop_name) = self.color_picker_prop {
                        self.editing = Some(EditingProp {
                            key: prop_name.clone(),
                            value: self.color_picker.color.to_hex(),
                            cursor: 7,
                        });
                        self.pending_commit = true;
                    }
                    self.color_picker_prop = None;
                    return;
                }
                ColorPickerEvent::None => { return; }
            }
        }

        // Font picker popup is open: route click to it
        if self.font_picker.open {
            let popup_x = rect.x + 10.0;
            let popup_y = rect.y + HEADER_H + TAB_H + 40.0;
            match self.font_picker.handle_click(mx, my, popup_x, popup_y) {
                FontPickerEvent::Changed { family, size } => {
                    self.editing = Some(EditingProp {
                        key: "Font".to_string(),
                        value: format!("{}, {}px", family, size),
                        cursor: 0,
                    });
                    return;
                }
                FontPickerEvent::Closed => {
                    self.editing = Some(EditingProp {
                        key: "Font".to_string(),
                        value: self.font_picker.to_string(),
                        cursor: 0,
                    });
                    self.pending_commit = true;
                    self.font_picker_active = false;
                    return;
                }
                FontPickerEvent::None => { return; }
            }
        }

        // Dropdown popup overlay routing
        if let Some(ref mut dropdown) = self.dropdown {
            let (px, py) = self.dropdown_pos.unwrap_or((rect.x + 10.0, rect.y + HEADER_H + TAB_H + 40.0));
            match dropdown.handle_mouse_at(mx, my, px, py, true) {
                DropdownEvent::Selected(idx) => {
                    if let Some(ref prop_name) = self.dropdown_prop {
                        self.editing = Some(EditingProp {
                            key: prop_name.clone(),
                            value: dropdown.items[idx].clone(),
                            cursor: dropdown.items[idx].len(),
                        });
                        self.pending_commit = true;
                    }
                    self.dropdown = None;
                    self.dropdown_prop = None;
                    self.dropdown_pos = None;
                    return;
                }
                DropdownEvent::Closed => {
                    self.dropdown = None;
                    self.dropdown_prop = None;
                    self.dropdown_pos = None;
                    return;
                }
                DropdownEvent::None => { return; }
            }
        }

        // Tab switching
        let tab_y = rect.y + HEADER_H;
        if my >= tab_y && my < tab_y + TAB_H {
            let tab_w = (rect.w - 2.0) / 2.0;
            if mx < rect.x + 1.0 + tab_w {
                self.tab = PropTab::Properties;
            } else {
                self.tab = PropTab::Events;
            }
            self.scroll_y = 0.0;
            self.editing = None;
            return;
        }

        let content_top = tab_y + TAB_H;
        if my < content_top { return; }

        let items = match self.tab {
            PropTab::Properties => self.collect_props(form, selected_control),
            PropTab::Events => self.collect_events(form, selected_control),
        };

        let val_x = rect.x + rect.w * 0.42;
        let mut y = content_top + 2.0 - self.scroll_y;

        for item in &items {
            match item {
                PropItem::Section(_) => { y += SECTION_H; }
                PropItem::Row(key, value) => {
                    if my >= y && my < y + ROW_H {
                        if self.tab == PropTab::Properties && mx >= val_x {
                            if key == "Type" {
                                // Type is read-only
                            } else if key == "BackColor" || key == "ForeColor" {
                                self.color_picker.set_from_hex(value);
                                self.color_picker.open = true;
                                self.color_picker_prop = Some(key.clone());
                                self.editing = None;
                            } else if key == "Font" {
                                self.font_picker.set_from_string(value);
                                self.font_picker.open = true;
                                self.font_picker_active = true;
                                self.editing = None;
                            } else if key.contains("Build ConnStr") || key.contains("Test Connection") {
                                if key.contains("Build") {
                                    self.conn_builder_open = true;
                                    self.wizard_props.clear();
                                    if let (Some(f), Some(c_name)) = (form, selected_control) {
                                        if let Some(ctrl) = f.controls.iter().find(|c| c.name == c_name) {
                                            let db_type = Self::get_prop(ctrl, "DbType", "PostgreSQL");
                                            for p in ["DbHost", "DbPort", "DbName", "DbUser", "DbPassword", "DbPath"] {
                                                self.wizard_props.insert(p.to_string(), Self::get_prop(ctrl, p, ""));
                                            }
                                            if self.wizard_props.get("DbHost").map(String::as_str).unwrap_or("") == "" {
                                                self.wizard_props.insert("DbHost".to_string(), "localhost".to_string());
                                            }
                                            if self.wizard_props.get("DbPort").map(String::as_str).unwrap_or("") == "" {
                                                let p = if db_type == "MySQL" { "3306" } else { "5432" };
                                                self.wizard_props.insert("DbPort".to_string(), p.to_string());
                                            }
                                            if self.wizard_props.get("DbUser").map(String::as_str).unwrap_or("") == "" {
                                                let u = if db_type == "MySQL" { "root" } else { "postgres" };
                                                self.wizard_props.insert("DbUser".to_string(), u.to_string());
                                            }
                                            self.wizard_props.insert("DbType".to_string(), db_type);
                                        }
                                    }
                                } else {
                                    self.pending_action = Some("test_conn".to_string());
                                }
                                self.editing = None;
                            } else {
                                self.editing = Some(EditingProp {
                                    key: key.clone(),
                                    value: value.clone(),
                                    cursor: value.len(),
                                });
                            }
                        } else if self.tab == PropTab::Events {
                            let ctrl_name = selected_control
                                .map(|s| s.to_string())
                                .or_else(|| form.map(|f| f.name.clone()))
                                .unwrap_or_default();
                            self.pending_event = Some((ctrl_name, key.clone()));
                            self.editing = None;
                        }
                        return;
                    }
                    y += ROW_H;
                }
                PropItem::CheckboxRow(key, checked) => {
                    if my >= y && my < y + ROW_H {
                        if self.tab == PropTab::Properties && mx >= val_x {
                            let new_val = if *checked { "False" } else { "True" };
                            self.editing = Some(EditingProp {
                                key: key.clone(),
                                value: new_val.to_string(),
                                cursor: new_val.len(),
                            });
                            self.pending_commit = true;
                        } else if self.tab == PropTab::Events {
                            let ctrl_name = selected_control.map(|s| s.to_string()).or_else(|| form.map(|f| f.name.clone())).unwrap_or_default();
                            self.pending_event = Some((ctrl_name, key.clone()));
                            self.editing = None;
                        }
                        return;
                    }
                    y += ROW_H;
                }
                PropItem::DropdownRow(key, current, options) => {
                    if my >= y && my < y + ROW_H {
                        if self.tab == PropTab::Properties && mx >= val_x {
                            let curr_idx = options.iter().position(|o| o == current).unwrap_or(0);
                            self.dropdown = Some(Dropdown::new(options.clone(), curr_idx, scale, Some(1)));
                            self.dropdown_prop = Some(key.clone());
                            self.dropdown_pos = Some((val_x, y + ROW_H));
                        }
                        return;
                    }
                    y += ROW_H;
                }
            }
        }

        self.editing = None;
    }

    pub fn scroll(&mut self, delta: f32, rect: LayoutRect, form: Option<&Form>, selected_control: Option<&str>) {
        let items = match self.tab {
            PropTab::Properties => self.collect_props(form, selected_control),
            PropTab::Events => self.collect_events(form, selected_control),
        };
        let total_h = self.items_height(&items);
        let content_h = rect.h - HEADER_H - TAB_H;
        let max_scroll = (total_h - content_h).max(0.0);
        self.scroll_y = (self.scroll_y - delta * ROW_H * 3.0).clamp(0.0, max_scroll);
    }

    pub fn handle_key(&mut self, key: &str) -> bool {
        let Some(ed) = &mut self.editing else { return false; };
        match key {
            "Left" => {
                if ed.cursor > 0 {
                    while ed.cursor > 0 {
                        ed.cursor -= 1;
                        if ed.value.is_char_boundary(ed.cursor) { break; }
                    }
                }
            }
            "Right" => {
                if ed.cursor < ed.value.len() {
                    while ed.cursor < ed.value.len() {
                        ed.cursor += 1;
                        if ed.value.is_char_boundary(ed.cursor) { break; }
                    }
                }
            }
            "Home" => { ed.cursor = 0; }
            "End" => { ed.cursor = ed.value.len(); }
            "Backspace" => {
                if ed.cursor > 0 {
                    let prev = ed.cursor;
                    while ed.cursor > 0 {
                        ed.cursor -= 1;
                        if ed.value.is_char_boundary(ed.cursor) { break; }
                    }
                    ed.value.drain(ed.cursor..prev);
                }
            }
            "Delete" => {
                if ed.cursor < ed.value.len() {
                    let prev = ed.cursor;
                    let mut next = ed.cursor;
                    while next < ed.value.len() {
                        next += 1;
                        if ed.value.is_char_boundary(next) { break; }
                    }
                    ed.value.drain(prev..next);
                }
            }
            "Enter" | "Tab" => {
                if key == "Tab" && self.conn_builder_open {
                    let mut handled = false;
                    if let Some(ed) = &self.editing {
                        let db_type = self.wizard_props.get("DbType").map(String::as_str).unwrap_or("PostgreSQL");
                        let fields = if db_type == "SQLite" {
                            vec!["DbPath"]
                        } else {
                            vec!["DbHost", "DbPort", "DbName", "DbUser", "DbPassword"]
                        };

                        if ed.key.starts_with("Wizard_") {
                            let k = ed.key.replace("Wizard_", "");
                            let mut next_idx = 0;
                            if let Some(pos) = fields.iter().position(|f| *f == k.as_str()) {
                                next_idx = (pos + 1) % fields.len();
                            }

                            self.wizard_props.insert(k, ed.value.clone());

                            let next_key = fields[next_idx];
                            let next_val = self.wizard_props.get(next_key).map(String::as_str).unwrap_or("").to_string();
                            self.editing = Some(EditingProp {
                                key: format!("Wizard_{}", next_key),
                                value: next_val.clone(),
                                cursor: next_val.len(),
                            });
                            handled = true;
                        }
                    }
                    if handled { return false; }
                }
                return true;
            }
            "Escape" => { return true; }
            _ => { return false; }
        }
        false
    }

    pub fn handle_char(&mut self, ch: char) -> bool {
        let Some(ed) = &mut self.editing else { return false; };
        if ch.is_control() { return false; }
        ed.value.insert(ed.cursor, ch);
        ed.cursor += ch.len_utf8();
        true
    }

    pub fn commit_edit(&mut self) -> Option<(String, String)> {
        self.editing.take().map(|ed| (ed.key, ed.value))
    }
}

fn fill(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pix.fill_rect(r, paint, Transform::identity(), None);
    }
}

fn stroke_rect(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    let mut pb = tiny_skia::PathBuilder::new();
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pb.push_rect(r);
    }
    if let Some(path) = pb.finish() {
        let mut st = tiny_skia::Stroke::default();
        st.width = 1.0 * s;
        pix.stroke_path(&path, paint, &st, Transform::identity(), None);
    }
}
