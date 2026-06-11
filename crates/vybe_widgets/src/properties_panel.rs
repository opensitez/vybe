//! Generic properties panel widget.
//!
//! Consumers build a `Vec<PropItem>` describing the current selection and
//! pass it to `draw`/`handle_click`/`handle_key`. The widget owns only UI
//! state (scroll, active tab, color picker, `TextInput` editor, open
//! `Dropdown`, `ScrollBar`) — not the underlying data model.

use crate::color_picker::{ColorPicker, ColorPickerEvent};
use crate::ide_text::draw_text;
use crate::{
    Checkbox, Dropdown, DropdownEvent, FontSystem, KeyEvent, PanelWidget, ScrollBar, SwashCache,
    TextColor as CosmicColor, TextInput,
    layout::{
        CheckState, LayoutRect, MouseButton, MouseEvent, MouseEventKind, RenderContext, WidgetEvent,
    },
};
use tiny_skia::{Paint, PathBuilder, Pixmap, Stroke, Transform};
use winit::keyboard::{Key, NamedKey};

#[derive(Clone)]
pub enum PropItem {
    Section(String),
    Row(String, String),
    CheckboxRow(String, bool),
    DropdownRow(String, String, Vec<String>),
    /// Multi-line list editor (e.g. `Items`). Opens an overlay on click.
    /// Stored value is `\n`-separated.
    MultilineRow(String, String),
}

#[derive(Clone, Copy, PartialEq)]
pub enum PropTab {
    Properties,
    Events,
}

/// In-progress inline edit: the property key plus the focused TextInput.
pub struct EditingProp {
    pub key: String,
    pub input: TextInput,
}

/// Popped-open dropdown: key, anchor x/y, and the `Dropdown` widget.
pub struct OpenDropdown {
    pub key: String,
    pub x: f32,
    pub y: f32,
    pub dropdown: Dropdown,
}

/// Event produced by panel interactions.
pub enum PropEvent {
    None,
    TabChanged(PropTab),
    ColorPickerChanged {
        prop_name: String,
        hex: String,
    },
    ColorPickerClosed {
        prop_name: String,
        hex: String,
    },
    /// User pressed Enter (or clicked away) after editing a value.
    ValueCommitted {
        key: String,
        value: String,
    },
    /// Checkbox clicked — caller should flip the backing property.
    ValueToggled {
        key: String,
        value: bool,
    },
    /// Dropdown item selected — caller should persist the new value.
    ValueSelected {
        key: String,
        value: String,
    },
    /// User clicked an event row in the Events tab. Caller should insert /
    /// navigate to the matching handler in code-behind.
    EventHandlerRequested {
        event: String,
    },
}

pub const PROP_HEADER_H: f32 = 28.0;
pub const PROP_TAB_H: f32 = 26.0;
pub const PROP_ROW_H: f32 = 24.0;
pub const PROP_SECTION_H: f32 = 20.0;
pub const PROP_SCROLLBAR_W: f32 = 10.0;

pub struct PropertiesPanel {
    pub tab: PropTab,
    pub scroll_y: f32,
    pub color_picker: ColorPicker,
    pub color_picker_prop: Option<String>,
    pub editing: Option<EditingProp>,
    pub dropdown: Option<OpenDropdown>,
    pub scrollbar: ScrollBar,
}

impl Default for PropertiesPanel {
    fn default() -> Self {
        Self {
            tab: PropTab::Properties,
            scroll_y: 0.0,
            color_picker: ColorPicker::new(),
            color_picker_prop: None,
            editing: None,
            dropdown: None,
            scrollbar: ScrollBar::new(true),
        }
    }
}

impl PropertiesPanel {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn items_height(items: &[PropItem]) -> f32 {
        items
            .iter()
            .map(|i| match i {
                PropItem::Section(_) => PROP_SECTION_H,
                PropItem::Row(..)
                | PropItem::CheckboxRow(..)
                | PropItem::DropdownRow(..)
                | PropItem::MultilineRow(..) => PROP_ROW_H,
            })
            .sum()
    }

    pub fn is_editing(&self) -> bool {
        self.editing.is_some()
    }

    /// End any in-progress edit and emit a commit event if there was one.
    pub fn commit_edit(&mut self) -> PropEvent {
        if let Some(edit) = self.editing.take() {
            PropEvent::ValueCommitted {
                key: edit.key,
                value: edit.input.value,
            }
        } else {
            PropEvent::None
        }
    }

    fn start_edit(
        &mut self,
        key: &str,
        current_value: &str,
        panel_x: f32,
        panel_w: f32,
        row_y: f32,
    ) {
        let val_x = panel_x + panel_w * 0.42;
        let val_w = panel_w - (val_x - panel_x) - PROP_SCROLLBAR_W - 2.0;
        let mut input = TextInput::new().with_name("prop_edit");
        input.value = current_value.to_string();
        input.cursor = current_value.len();
        input.set_rect(LayoutRect {
            x: val_x,
            y: row_y,
            w: val_w,
            h: PROP_ROW_H,
        });
        input.set_focused(true);
        self.editing = Some(EditingProp {
            key: key.to_string(),
            input,
        });
    }

    /// Draw the panel. Call this once per frame.
    pub fn draw(
        &mut self,
        ctx: &mut RenderContext,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        items: &[PropItem],
    ) {
        let s = ctx.scale;
        let mut paint = Paint::default();

        // Background
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(ctx.pixmap, &paint, x, y, w, h, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(ctx.pixmap, &paint, x, y, 1.0, h, s);

        // Title + tabs
        let title_color = CosmicColor::rgba(50, 50, 50, 255);
        draw_text(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            "Properties",
            x + 10.0,
            y + 6.0,
            13.0,
            title_color,
            s,
        );
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(ctx.pixmap, &paint, x, y + PROP_HEADER_H - 1.0, w, 1.0, s);

        let tab_y = y + PROP_HEADER_H;
        let tab_w = (w - 2.0) / 2.0;
        let text_color = CosmicColor::rgba(30, 30, 30, 255);
        let dim_color = CosmicColor::rgba(100, 100, 100, 255);

        draw_tab(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            "Properties",
            x + 1.0,
            tab_y,
            tab_w,
            self.tab == PropTab::Properties,
            text_color,
            dim_color,
            s,
        );
        draw_tab(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            "Events",
            x + 1.0 + tab_w,
            tab_y,
            tab_w,
            self.tab == PropTab::Events,
            text_color,
            dim_color,
            s,
        );

        paint.set_color_rgba8(204, 204, 204, 255);
        fill(ctx.pixmap, &paint, x, tab_y + PROP_TAB_H - 1.0, w, 1.0, s);

        // Content
        let content_top = tab_y + PROP_TAB_H;
        let content_h = h - PROP_HEADER_H - PROP_TAB_H;

        let label_color = CosmicColor::rgba(80, 80, 80, 255);
        let val_color = CosmicColor::rgba(30, 30, 30, 255);
        let section_color = CosmicColor::rgba(100, 100, 100, 255);
        let val_x = x + w * 0.42;
        let mut row_y = content_top + 2.0 - self.scroll_y;
        let mut row_idx = 0usize;
        let mut editor_row_y: Option<f32> = None;

        for item in items {
            match item {
                PropItem::Section(label) => {
                    if row_y + PROP_SECTION_H > content_top && row_y < content_top + content_h {
                        paint.set_color_rgba8(240, 240, 240, 255);
                        fill(
                            ctx.pixmap,
                            &paint,
                            x + 1.0,
                            row_y,
                            w - 2.0,
                            PROP_SECTION_H,
                            s,
                        );
                        draw_text(
                            ctx.pixmap,
                            ctx.font_system,
                            ctx.swash_cache,
                            label,
                            x + 8.0,
                            row_y + 3.0,
                            10.0,
                            section_color,
                            s,
                        );
                    }
                    row_y += PROP_SECTION_H;
                }
                PropItem::Row(key, value) => {
                    let being_edited = self
                        .editing
                        .as_ref()
                        .map(|e| &e.key == key)
                        .unwrap_or(false);
                    if row_y + PROP_ROW_H > content_top && row_y < content_top + content_h {
                        if being_edited {
                            paint.set_color_rgba8(227, 242, 253, 255);
                            fill(ctx.pixmap, &paint, x + 1.0, row_y, w - 2.0, PROP_ROW_H, s);
                        } else if row_idx % 2 == 1 {
                            paint.set_color_rgba8(247, 247, 247, 255);
                            fill(ctx.pixmap, &paint, x + 1.0, row_y, w - 2.0, PROP_ROW_H, s);
                        }
                        draw_text(
                            ctx.pixmap,
                            ctx.font_system,
                            ctx.swash_cache,
                            key,
                            x + 8.0,
                            row_y + 4.0,
                            11.0,
                            label_color,
                            s,
                        );
                        paint.set_color_rgba8(230, 230, 230, 255);
                        fill(ctx.pixmap, &paint, val_x - 2.0, row_y, 1.0, PROP_ROW_H, s);

                        if !being_edited {
                            draw_text(
                                ctx.pixmap,
                                ctx.font_system,
                                ctx.swash_cache,
                                value,
                                val_x + 4.0,
                                row_y + 4.0,
                                11.0,
                                val_color,
                                s,
                            );

                            if key == "BackColor" || key == "ForeColor" {
                                let sw_x = x + w - PROP_SCROLLBAR_W - 22.0;
                                let sw_y = row_y + 3.0;
                                let sw_sz = PROP_ROW_H - 6.0;
                                if let Some(c) = crate::color_picker::PickedColor::from_hex(value) {
                                    let mut sp = Paint::default();
                                    sp.set_color_rgba8(c.r, c.g, c.b, c.a);
                                    fill(ctx.pixmap, &sp, sw_x, sw_y, sw_sz, sw_sz, s);
                                    sp.set_color_rgba8(160, 160, 160, 255);
                                    stroke_rect(ctx.pixmap, &sp, sw_x, sw_y, sw_sz, sw_sz, s);
                                }
                            }
                        } else {
                            editor_row_y = Some(row_y);
                        }

                        paint.set_color_rgba8(235, 235, 235, 255);
                        fill(
                            ctx.pixmap,
                            &paint,
                            x + 1.0,
                            row_y + PROP_ROW_H - 1.0,
                            w - 2.0,
                            1.0,
                            s,
                        );
                    }
                    row_idx += 1;
                    row_y += PROP_ROW_H;
                }
                PropItem::CheckboxRow(key, checked) => {
                    if row_y + PROP_ROW_H > content_top && row_y < content_top + content_h {
                        if row_idx % 2 == 1 {
                            paint.set_color_rgba8(247, 247, 247, 255);
                            fill(ctx.pixmap, &paint, x + 1.0, row_y, w - 2.0, PROP_ROW_H, s);
                        }
                        draw_text(
                            ctx.pixmap,
                            ctx.font_system,
                            ctx.swash_cache,
                            key,
                            x + 8.0,
                            row_y + 4.0,
                            11.0,
                            label_color,
                            s,
                        );
                        paint.set_color_rgba8(230, 230, 230, 255);
                        fill(ctx.pixmap, &paint, val_x - 2.0, row_y, 1.0, PROP_ROW_H, s);

                        // Real Checkbox widget: paint at value column.
                        let cb_x = val_x + 4.0;
                        let cb_y = row_y + 4.0;
                        let mut cb = Checkbox::new("");
                        cb.size = 14.0;
                        cb.check_state = if *checked {
                            CheckState::Checked
                        } else {
                            CheckState::Unchecked
                        };
                        cb.paint(ctx.pixmap, cb_x, cb_y, s);

                        let lbl = if *checked { "True" } else { "False" };
                        draw_text(
                            ctx.pixmap,
                            ctx.font_system,
                            ctx.swash_cache,
                            lbl,
                            cb_x + cb.size + 4.0,
                            row_y + 4.0,
                            11.0,
                            val_color,
                            s,
                        );
                        paint.set_color_rgba8(235, 235, 235, 255);
                        fill(
                            ctx.pixmap,
                            &paint,
                            x + 1.0,
                            row_y + PROP_ROW_H - 1.0,
                            w - 2.0,
                            1.0,
                            s,
                        );
                    }
                    let _ = key;
                    row_idx += 1;
                    row_y += PROP_ROW_H;
                }
                PropItem::MultilineRow(key, value) => {
                    let being_edited = self
                        .editing
                        .as_ref()
                        .map(|e| &e.key == key)
                        .unwrap_or(false);
                    if row_y + PROP_ROW_H > content_top && row_y < content_top + content_h {
                        if being_edited {
                            paint.set_color_rgba8(227, 242, 253, 255);
                            fill(ctx.pixmap, &paint, x + 1.0, row_y, w - 2.0, PROP_ROW_H, s);
                        } else if row_idx % 2 == 1 {
                            paint.set_color_rgba8(247, 247, 247, 255);
                            fill(ctx.pixmap, &paint, x + 1.0, row_y, w - 2.0, PROP_ROW_H, s);
                        }
                        draw_text(
                            ctx.pixmap,
                            ctx.font_system,
                            ctx.swash_cache,
                            key,
                            x + 8.0,
                            row_y + 4.0,
                            11.0,
                            label_color,
                            s,
                        );
                        paint.set_color_rgba8(230, 230, 230, 255);
                        fill(ctx.pixmap, &paint, val_x - 2.0, row_y, 1.0, PROP_ROW_H, s);
                        if !being_edited {
                            let joined = value.replace('\n', ", ");
                            let shown = if joined.is_empty() {
                                "(empty) …".to_string()
                            } else {
                                format!("{} …", joined)
                            };
                            draw_text(
                                ctx.pixmap,
                                ctx.font_system,
                                ctx.swash_cache,
                                &shown,
                                val_x + 4.0,
                                row_y + 4.0,
                                11.0,
                                val_color,
                                s,
                            );
                        } else {
                            editor_row_y = Some(row_y);
                        }
                        paint.set_color_rgba8(235, 235, 235, 255);
                        fill(
                            ctx.pixmap,
                            &paint,
                            x + 1.0,
                            row_y + PROP_ROW_H - 1.0,
                            w - 2.0,
                            1.0,
                            s,
                        );
                    }
                    row_idx += 1;
                    row_y += PROP_ROW_H;
                }
                PropItem::DropdownRow(key, current, _options) => {
                    if row_y + PROP_ROW_H > content_top && row_y < content_top + content_h {
                        if row_idx % 2 == 1 {
                            paint.set_color_rgba8(247, 247, 247, 255);
                            fill(ctx.pixmap, &paint, x + 1.0, row_y, w - 2.0, PROP_ROW_H, s);
                        }
                        draw_text(
                            ctx.pixmap,
                            ctx.font_system,
                            ctx.swash_cache,
                            key,
                            x + 8.0,
                            row_y + 4.0,
                            11.0,
                            label_color,
                            s,
                        );
                        paint.set_color_rgba8(230, 230, 230, 255);
                        fill(ctx.pixmap, &paint, val_x - 2.0, row_y, 1.0, PROP_ROW_H, s);
                        let dd_w = w - (val_x - x) - PROP_SCROLLBAR_W - 2.0;
                        paint.set_color_rgba8(255, 255, 255, 255);
                        fill(
                            ctx.pixmap,
                            &paint,
                            val_x,
                            row_y + 1.0,
                            dd_w,
                            PROP_ROW_H - 2.0,
                            s,
                        );
                        paint.set_color_rgba8(180, 180, 180, 255);
                        stroke_rect(
                            ctx.pixmap,
                            &paint,
                            val_x,
                            row_y + 1.0,
                            dd_w,
                            PROP_ROW_H - 2.0,
                            s,
                        );
                        draw_text(
                            ctx.pixmap,
                            ctx.font_system,
                            ctx.swash_cache,
                            current,
                            val_x + 4.0,
                            row_y + 4.0,
                            11.0,
                            val_color,
                            s,
                        );
                        let arr_x = val_x + dd_w - 14.0;
                        let arr_y = row_y + PROP_ROW_H / 2.0 - 2.0;
                        paint.set_color_rgba8(80, 80, 80, 255);
                        fill(ctx.pixmap, &paint, arr_x, arr_y, 8.0, 1.0, s);
                        fill(ctx.pixmap, &paint, arr_x + 1.0, arr_y + 1.0, 6.0, 1.0, s);
                        fill(ctx.pixmap, &paint, arr_x + 2.0, arr_y + 2.0, 4.0, 1.0, s);
                        fill(ctx.pixmap, &paint, arr_x + 3.0, arr_y + 3.0, 2.0, 1.0, s);
                        paint.set_color_rgba8(235, 235, 235, 255);
                        fill(
                            ctx.pixmap,
                            &paint,
                            x + 1.0,
                            row_y + PROP_ROW_H - 1.0,
                            w - 2.0,
                            1.0,
                            s,
                        );
                    }
                    let _ = key;
                    row_idx += 1;
                    row_y += PROP_ROW_H;
                }
            }
        }

        if items.is_empty() {
            draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                "No selection",
                x + 10.0,
                content_top + 8.0,
                12.0,
                dim_color,
                s,
            );
        }

        // Overdraw header/tabs to clip scrolled items
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(ctx.pixmap, &paint, x, y, w, PROP_HEADER_H + PROP_TAB_H, s);
        draw_text(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            "Properties",
            x + 10.0,
            y + 6.0,
            13.0,
            title_color,
            s,
        );
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(ctx.pixmap, &paint, x, y + PROP_HEADER_H - 1.0, w, 1.0, s);
        draw_tab(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            "Properties",
            x + 1.0,
            tab_y,
            tab_w,
            self.tab == PropTab::Properties,
            text_color,
            dim_color,
            s,
        );
        draw_tab(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            "Events",
            x + 1.0 + tab_w,
            tab_y,
            tab_w,
            self.tab == PropTab::Events,
            text_color,
            dim_color,
            s,
        );
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(ctx.pixmap, &paint, x, tab_y + PROP_TAB_H - 1.0, w, 1.0, s);

        // Scrollbar (real widget)
        let total_h = Self::items_height(items);
        if total_h > content_h {
            self.scrollbar.content_size = total_h;
            self.scrollbar.viewport_size = content_h;
            let max_scroll = (total_h - content_h).max(1.0);
            self.scrollbar.pos = (self.scroll_y / max_scroll).clamp(0.0, 1.0);
            self.scrollbar.set_rect(LayoutRect {
                x: x + w - PROP_SCROLLBAR_W,
                y: content_top,
                w: PROP_SCROLLBAR_W,
                h: content_h,
            });
            self.scrollbar
                .paint(ctx.pixmap, x + w - PROP_SCROLLBAR_W, content_top, s);
        }

        // Left border on top
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(ctx.pixmap, &paint, x, y, 1.0, h, s);

        // Render the TextInput editor last so it paints above the row.
        if let (Some(edit), Some(ry)) = (self.editing.as_mut(), editor_row_y) {
            let val_w = w - (val_x - x) - PROP_SCROLLBAR_W - 2.0;
            edit.input.set_rect(LayoutRect {
                x: val_x,
                y: ry,
                w: val_w,
                h: PROP_ROW_H,
            });
            edit.input.render(ctx);
        }

        // Open Dropdown popup on top of rows.
        if let Some(open) = self.dropdown.as_mut() {
            open.dropdown.scale = ctx.scale;
            open.dropdown.render_list(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                open.x,
                open.y,
                (255, 255, 255, 255),
                (180, 180, 180, 255),
                (227, 242, 253, 255),
                (240, 240, 240, 255),
                CosmicColor::rgba(30, 30, 30, 255),
                CosmicColor::rgba(80, 80, 80, 255),
            );
        }

        // Color picker popup on top of everything
        if self.color_picker.open {
            let popup_x = x + 10.0;
            let popup_y = y + PROP_HEADER_H + PROP_TAB_H + 40.0;
            self.color_picker
                .render_popup(ctx.pixmap, popup_x, popup_y, ctx.scale);
        }
    }

    /// Handle a mouse click. Returns a PropEvent the caller should react to.
    pub fn handle_click(
        &mut self,
        mx: f32,
        my: f32,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        items: &[PropItem],
    ) -> PropEvent {
        if mx < x || mx >= x + w || my < y || my >= y + h {
            return PropEvent::None;
        }

        // Open dropdown takes priority — click inside commits a selection,
        // click outside closes without selection.
        if let Some(mut open) = self.dropdown.take() {
            let evt = open.dropdown.handle_mouse_at(mx, my, open.x, open.y, true);
            match evt {
                DropdownEvent::Selected(idx) => {
                    let value = open.dropdown.items.get(idx).cloned().unwrap_or_default();
                    return PropEvent::ValueSelected {
                        key: open.key,
                        value,
                    };
                }
                DropdownEvent::Closed | DropdownEvent::None => {
                    // Closed — fall through so the click also counts normally.
                }
            }
        }

        // Click inside the active TextInput rect → forward to it for cursor
        // positioning; don't commit.
        if let Some(edit) = self.editing.as_mut() {
            if edit.input.rect().contains(mx, my) {
                let me = MouseEvent {
                    x: mx,
                    y: my,
                    kind: MouseEventKind::Press(MouseButton::Left),
                    cmd: false,
                    shift: false,
                    alt: false,
                };
                edit.input.handle_mouse(&me);
                edit.input.set_focused(true);
                return PropEvent::None;
            }
        }

        // Click in the scrollbar — only route there when the bar is actually
        // visible (content exceeds viewport), otherwise the right-edge of
        // the value column would eat clicks meant for a Row.
        let content_top = y + PROP_HEADER_H + PROP_TAB_H;
        let content_h = h - PROP_HEADER_H - PROP_TAB_H;
        let total_h = Self::items_height(items);
        if total_h > content_h
            && mx >= x + w - PROP_SCROLLBAR_W
            && my >= content_top
            && my < content_top + content_h
        {
            let sb_x = x + w - PROP_SCROLLBAR_W;
            if self.scrollbar.mouse_down(mx - sb_x, my - content_top) {
                self.sync_scroll_from_bar(content_h);
                return self.commit_edit();
            }
        }

        // Click landed outside the active input/dropdown/scrollbar —
        // commit edits and continue routing.
        let committed = self.commit_edit();

        // Color picker
        if self.color_picker.open {
            let popup_x = x + 10.0;
            let popup_y = y + PROP_HEADER_H + PROP_TAB_H + 40.0;
            match self.color_picker.handle_click(mx, my, popup_x, popup_y) {
                ColorPickerEvent::Changed(c) => {
                    let hex = c.to_hex();
                    if let Some(pn) = self.color_picker_prop.clone() {
                        return PropEvent::ColorPickerChanged { prop_name: pn, hex };
                    }
                    return committed;
                }
                ColorPickerEvent::Closed => {
                    let hex = self.color_picker.color.to_hex();
                    if let Some(pn) = self.color_picker_prop.take() {
                        return PropEvent::ColorPickerClosed { prop_name: pn, hex };
                    }
                    return committed;
                }
                ColorPickerEvent::None => return committed,
            }
        }

        // Tab switch
        let tab_y = y + PROP_HEADER_H;
        if my >= tab_y && my < tab_y + PROP_TAB_H {
            let tab_w = (w - 2.0) / 2.0;
            let new_tab = if mx < x + 1.0 + tab_w {
                PropTab::Properties
            } else {
                PropTab::Events
            };
            if new_tab != self.tab {
                self.tab = new_tab;
                self.scroll_y = 0.0;
                return PropEvent::TabChanged(new_tab);
            }
            return committed;
        }

        // Events tab: any click on a Row fires EventHandlerRequested.
        if self.tab == PropTab::Events && my >= content_top {
            let mut row_y = content_top + 2.0 - self.scroll_y;
            for item in items {
                let row_h = match item {
                    PropItem::Section(_) => PROP_SECTION_H,
                    _ => PROP_ROW_H,
                };
                if let PropItem::Row(key, _) = item {
                    if my >= row_y && my < row_y + PROP_ROW_H {
                        return PropEvent::EventHandlerRequested { event: key.clone() };
                    }
                }
                row_y += row_h;
            }
            return committed;
        }

        // Row click on value column
        if self.tab == PropTab::Properties {
            let val_x = x + w * 0.42;
            if my >= content_top && mx >= val_x {
                let mut row_y = content_top + 2.0 - self.scroll_y;
                for item in items {
                    match item {
                        PropItem::Section(_) => {
                            row_y += PROP_SECTION_H;
                        }
                        PropItem::Row(key, value) => {
                            if my >= row_y && my < row_y + PROP_ROW_H {
                                if key == "BackColor" || key == "ForeColor" {
                                    self.color_picker.set_from_hex(value);
                                    self.color_picker.open = true;
                                    self.color_picker_prop = Some(key.clone());
                                } else if !is_read_only(key) {
                                    self.start_edit(key, value, x, w, row_y);
                                }
                                return committed;
                            }
                            row_y += PROP_ROW_H;
                        }
                        PropItem::CheckboxRow(key, checked) => {
                            if my >= row_y && my < row_y + PROP_ROW_H {
                                let cb_x = val_x + 4.0;
                                let cb_y = row_y + 4.0;
                                let mut cb = Checkbox::new("");
                                cb.size = 14.0;
                                cb.check_state = if *checked {
                                    CheckState::Checked
                                } else {
                                    CheckState::Unchecked
                                };
                                if cb.click(mx - cb_x, my - cb_y) {
                                    return PropEvent::ValueToggled {
                                        key: key.clone(),
                                        value: !*checked,
                                    };
                                }
                                // Click on the row but outside the checkbox — still flip.
                                return PropEvent::ValueToggled {
                                    key: key.clone(),
                                    value: !*checked,
                                };
                            }
                            row_y += PROP_ROW_H;
                        }
                        PropItem::DropdownRow(key, current, options) => {
                            if my >= row_y && my < row_y + PROP_ROW_H {
                                let selected_idx =
                                    options.iter().position(|o| o == current).unwrap_or(0);
                                let mut dd =
                                    Dropdown::new(options.clone(), selected_idx, 1.0, Some(1))
                                        .with_name(key);
                                let (_dd_w, _dd_h) = dd.get_size();
                                let anchor_x = val_x;
                                let anchor_y = row_y + PROP_ROW_H;
                                dd.hover_idx = Some(selected_idx);
                                self.dropdown = Some(OpenDropdown {
                                    key: key.clone(),
                                    x: anchor_x,
                                    y: anchor_y,
                                    dropdown: dd,
                                });
                                return committed;
                            }
                            row_y += PROP_ROW_H;
                        }
                        PropItem::MultilineRow(key, value) => {
                            if my >= row_y && my < row_y + PROP_ROW_H {
                                // Open the single-line TextInput editor
                                // pre-filled with value joined by ", ".
                                // The caller's apply_value_prop splits on
                                // either ',' or '\n'.
                                let joined = value.replace('\n', ", ");
                                self.start_edit(key, &joined, x, w, row_y);
                                return committed;
                            }
                            row_y += PROP_ROW_H;
                        }
                    }
                }
            }
        }

        committed
    }

    /// Route a mouse move. Forward to scrollbar drag / color picker drag.
    pub fn handle_mouse_move(
        &mut self,
        mx: f32,
        my: f32,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
    ) -> PropEvent {
        let _ = (x, w);
        let content_top = y + PROP_HEADER_H + PROP_TAB_H;
        let content_h = h - PROP_HEADER_H - PROP_TAB_H;
        if self.scrollbar.dragging {
            let sb_x = x + w - PROP_SCROLLBAR_W;
            self.scrollbar.mouse_move(mx - sb_x, my - content_top);
            self.sync_scroll_from_bar(content_h);
        }
        if let Some(open) = self.dropdown.as_mut() {
            // Hover update — no event emitted.
            let _ = open.dropdown.handle_mouse_at(mx, my, open.x, open.y, false);
        }
        PropEvent::None
    }

    pub fn handle_mouse_up(&mut self) {
        self.scrollbar.mouse_up();
    }

    fn sync_scroll_from_bar(&mut self, _content_h: f32) {
        self.scroll_y = self.scrollbar.scroll_offset();
    }

    /// Keyboard input. When an edit is active, forwards to the `TextInput`;
    /// Enter commits (returning `ValueCommitted`), Escape cancels.
    pub fn handle_key(&mut self, event: &KeyEvent) -> PropEvent {
        let Some(edit) = self.editing.as_mut() else {
            return PropEvent::None;
        };
        if event.state != winit::event::ElementState::Pressed {
            return PropEvent::None;
        }

        match &event.logical_key {
            Key::Named(NamedKey::Enter) => {
                let taken = self.editing.take().unwrap();
                return PropEvent::ValueCommitted {
                    key: taken.key,
                    value: taken.input.value,
                };
            }
            Key::Named(NamedKey::Escape) => {
                self.editing = None;
                return PropEvent::None;
            }
            _ => {}
        }

        let _ = edit.input.handle_key(event);
        let _: Vec<WidgetEvent> = edit.input.drain_events();
        PropEvent::None
    }

    pub fn scroll(&mut self, amount: f32, items: &[PropItem], visible_h: f32) {
        let total_h = Self::items_height(items);
        let max_scroll = (total_h - visible_h).max(0.0);
        self.scroll_y = (self.scroll_y - amount).clamp(0.0, max_scroll);
    }
}

fn draw_tab(
    pix: &mut Pixmap,
    fs: &mut FontSystem,
    sc: &mut SwashCache,
    label: &str,
    x: f32,
    y: f32,
    w: f32,
    active: bool,
    text_color: CosmicColor,
    dim_color: CosmicColor,
    s: f32,
) {
    let mut paint = Paint::default();
    if active {
        paint.set_color_rgba8(227, 242, 253, 255);
    } else {
        paint.set_color_rgba8(245, 245, 245, 255);
    }
    fill(pix, &paint, x, y, w, PROP_TAB_H, s);
    draw_text(
        pix,
        fs,
        sc,
        label,
        x + 13.0,
        y + 5.0,
        12.0,
        if active { text_color } else { dim_color },
        s,
    );
}

/// Rows whose value is derived / computed and can't be edited in-place.
fn is_read_only(key: &str) -> bool {
    matches!(key, "Type")
}

fn fill(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pix.fill_rect(r, paint, Transform::identity(), None);
    }
}

fn stroke_rect(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    let mut pb = PathBuilder::new();
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pb.push_rect(r);
    }
    if let Some(path) = pb.finish() {
        let mut stroke = Stroke::default();
        stroke.width = 1.0 * s;
        pix.stroke_path(&path, paint, &stroke, Transform::identity(), None);
    }
}
