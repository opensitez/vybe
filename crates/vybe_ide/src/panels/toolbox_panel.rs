//! Toolbox panel — control palette matching the legacy editor, with Data section.

use cosmic_text::{Color as CosmicColor, FontSystem, SwashCache};
use tiny_skia::{Paint, Pixmap, Transform};

use crate::layout::Rect;
use crate::text::draw_text;
use vybe_forms::ControlType;

#[derive(Debug, Clone, PartialEq)]
pub enum ControlTool {
    Pointer,
    Control(ControlType),
}

struct ToolEntry {
    label: &'static str,
    tool: ControlTool,
}

struct SectionHeader {
    label: &'static str,
}

enum ToolItem {
    Entry(ToolEntry),
    Header(SectionHeader),
}

fn build_items() -> Vec<ToolItem> {
    use ControlType::*;
    vec![
        // Standard controls
        ToolItem::Entry(ToolEntry { label: "Pointer", tool: ControlTool::Pointer }),
        ToolItem::Entry(ToolEntry { label: "Button", tool: ControlTool::Control(Button) }),
        ToolItem::Entry(ToolEntry { label: "Label", tool: ControlTool::Control(Label) }),
        ToolItem::Entry(ToolEntry { label: "TextBox", tool: ControlTool::Control(TextBox) }),
        ToolItem::Entry(ToolEntry { label: "CheckBox", tool: ControlTool::Control(CheckBox) }),
        ToolItem::Entry(ToolEntry { label: "RadioButton", tool: ControlTool::Control(RadioButton) }),
        ToolItem::Entry(ToolEntry { label: "ComboBox", tool: ControlTool::Control(ComboBox) }),
        ToolItem::Entry(ToolEntry { label: "ListBox", tool: ControlTool::Control(ListBox) }),
        ToolItem::Entry(ToolEntry { label: "GroupBox", tool: ControlTool::Control(Frame) }),
        ToolItem::Entry(ToolEntry { label: "PictureBox", tool: ControlTool::Control(PictureBox) }),
        ToolItem::Entry(ToolEntry { label: "RichTextBox", tool: ControlTool::Control(RichTextBox) }),
        ToolItem::Entry(ToolEntry { label: "WebBrowser", tool: ControlTool::Control(WebBrowser) }),
        ToolItem::Entry(ToolEntry { label: "TreeView", tool: ControlTool::Control(TreeView) }),
        ToolItem::Entry(ToolEntry { label: "DataGridView", tool: ControlTool::Control(DataGridView) }),
        ToolItem::Entry(ToolEntry { label: "Panel", tool: ControlTool::Control(Panel) }),
        ToolItem::Entry(ToolEntry { label: "ListView", tool: ControlTool::Control(ListView) }),
        ToolItem::Entry(ToolEntry { label: "TabControl", tool: ControlTool::Control(TabControl) }),
        ToolItem::Entry(ToolEntry { label: "ProgressBar", tool: ControlTool::Control(ProgressBar) }),
        ToolItem::Entry(ToolEntry { label: "NumericUpDown", tool: ControlTool::Control(NumericUpDown) }),
        ToolItem::Entry(ToolEntry { label: "MenuStrip", tool: ControlTool::Control(MenuStrip) }),
        ToolItem::Entry(ToolEntry { label: "ContextMenuStrip", tool: ControlTool::Control(ContextMenuStrip) }),
        ToolItem::Entry(ToolEntry { label: "StatusStrip", tool: ControlTool::Control(StatusStrip) }),
        ToolItem::Entry(ToolEntry { label: "DateTimePicker", tool: ControlTool::Control(DateTimePicker) }),
        ToolItem::Entry(ToolEntry { label: "LinkLabel", tool: ControlTool::Control(LinkLabel) }),
        ToolItem::Entry(ToolEntry { label: "ToolStrip", tool: ControlTool::Control(ToolStrip) }),
        ToolItem::Entry(ToolEntry { label: "TrackBar", tool: ControlTool::Control(TrackBar) }),
        ToolItem::Entry(ToolEntry { label: "MaskedTextBox", tool: ControlTool::Control(MaskedTextBox) }),
        ToolItem::Entry(ToolEntry { label: "SplitContainer", tool: ControlTool::Control(SplitContainer) }),
        ToolItem::Entry(ToolEntry { label: "FlowLayoutPanel", tool: ControlTool::Control(FlowLayoutPanel) }),
        ToolItem::Entry(ToolEntry { label: "TableLayoutPanel", tool: ControlTool::Control(TableLayoutPanel) }),
        ToolItem::Entry(ToolEntry { label: "MonthCalendar", tool: ControlTool::Control(MonthCalendar) }),
        ToolItem::Entry(ToolEntry { label: "HScrollBar", tool: ControlTool::Control(HScrollBar) }),
        ToolItem::Entry(ToolEntry { label: "VScrollBar", tool: ControlTool::Control(VScrollBar) }),
        // Data section
        ToolItem::Header(SectionHeader { label: "DATA" }),
        ToolItem::Entry(ToolEntry { label: "\u{1F517} BindingSource", tool: ControlTool::Control(BindingSourceComponent) }),
        ToolItem::Entry(ToolEntry { label: "\u{1F9ED} BindingNavigator", tool: ControlTool::Control(BindingNavigator) }),
        ToolItem::Entry(ToolEntry { label: "\u{1F5C4} DataSet", tool: ControlTool::Control(DataSetComponent) }),
        ToolItem::Entry(ToolEntry { label: "\u{1F4CB} DataTable", tool: ControlTool::Control(DataTableComponent) }),
        ToolItem::Entry(ToolEntry { label: "\u{1F50C} DataAdapter", tool: ControlTool::Control(DataAdapterComponent) }),
    ]
}

const ITEM_H: f32 = 22.0;
const HEADER_H: f32 = 26.0;
const SECTION_H: f32 = 20.0;
const SCROLLBAR_W: f32 = 10.0;

pub struct ToolboxPanel {
    pub selected_idx: Option<usize>, // index into the entry list (not items list)
    pub scroll_y: f32,
}

impl ToolboxPanel {
    pub fn new() -> Self {
        Self { selected_idx: Some(0), scroll_y: 0.0 }
    }

    pub fn selected_tool(&self) -> ControlTool {
        let items = build_items();
        let mut entry_idx = 0usize;
        for item in &items {
            if let ToolItem::Entry(e) = item {
                if self.selected_idx == Some(entry_idx) {
                    return e.tool.clone();
                }
                entry_idx += 1;
            }
        }
        ControlTool::Pointer
    }

    pub fn reset_to_pointer(&mut self) {
        self.selected_idx = Some(0);
    }

    fn content_h(&self) -> f32 {
        let items = build_items();
        let mut h = 0.0;
        for item in &items {
            match item {
                ToolItem::Entry(_) => h += ITEM_H,
                ToolItem::Header(_) => h += SECTION_H,
            }
        }
        h
    }

    fn max_scroll(&self, rect: Rect) -> f32 {
        (self.content_h() - (rect.h - HEADER_H)).max(0.0)
    }

    pub fn scroll(&mut self, delta: f32, rect: Rect) {
        self.scroll_y = (self.scroll_y - delta * ITEM_H * 3.0)
            .max(0.0)
            .min(self.max_scroll(rect));
    }

    pub fn render(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        rect: Rect, scale: f32,
    ) {
        let s = scale;
        let mut paint = Paint::default();

        // Background
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);

        // Title
        let title_color = CosmicColor::rgba(50, 50, 50, 255);
        draw_text(pix, fs, sc, "Toolbox", rect.x + 10.0, rect.y + 6.0, 13.0, title_color, s);

        // Separator
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + HEADER_H - 1.0, rect.w, 1.0, s);

        let items = build_items();
        let text_color = CosmicColor::rgba(30, 30, 30, 255);
        let section_color = CosmicColor::rgba(100, 100, 100, 255);
        let list_top = rect.y + HEADER_H;
        let list_h = rect.h - HEADER_H;
        let mut y = list_top - self.scroll_y;
        let mut entry_idx = 0usize;

        for item in &items {
            match item {
                ToolItem::Header(hdr) => {
                    if y + SECTION_H > list_top && y < list_top + list_h {
                        // Section separator line
                        paint.set_color_rgba8(204, 204, 204, 255);
                        fill(pix, &paint, rect.x + 8.0, y + 2.0, rect.w - 16.0, 1.0, s);
                        draw_text(pix, fs, sc, hdr.label, rect.x + 10.0, y + 6.0, 10.0, section_color, s);
                    }
                    y += SECTION_H;
                }
                ToolItem::Entry(e) => {
                    if y + ITEM_H > list_top && y < list_top + list_h {
                        let is_sel = self.selected_idx == Some(entry_idx);
                        if is_sel {
                            paint.set_color_rgba8(0, 120, 212, 255);
                            fill(pix, &paint, rect.x + 4.0, y, rect.w - SCROLLBAR_W - 8.0, ITEM_H, s);
                            let white = CosmicColor::rgba(255, 255, 255, 255);
                            draw_text(pix, fs, sc, e.label, rect.x + 12.0, y + 3.0, 12.0, white, s);
                        } else {
                            draw_text(pix, fs, sc, e.label, rect.x + 12.0, y + 3.0, 12.0, text_color, s);
                        }
                    }
                    entry_idx += 1;
                    y += ITEM_H;
                }
            }
        }

        // Overdraw header to clip
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, HEADER_H, s);
        draw_text(pix, fs, sc, "Toolbox", rect.x + 10.0, rect.y + 6.0, 13.0, title_color, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + HEADER_H - 1.0, rect.w, 1.0, s);

        // Scrollbar
        let max_scroll = self.max_scroll(rect);
        if max_scroll > 0.0 {
            let sb_x = rect.x + rect.w - SCROLLBAR_W;
            paint.set_color_rgba8(235, 235, 235, 255);
            fill(pix, &paint, sb_x, list_top, SCROLLBAR_W, list_h, s);
            let visible_frac = (list_h / self.content_h()).min(1.0);
            let thumb_h = (list_h * visible_frac).max(20.0);
            let scroll_frac = self.scroll_y / max_scroll;
            let thumb_y = list_top + scroll_frac * (list_h - thumb_h);
            paint.set_color_rgba8(190, 190, 190, 255);
            fill(pix, &paint, sb_x + 2.0, thumb_y, SCROLLBAR_W - 4.0, thumb_h, s);
        }

        // Right border
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x + rect.w - 1.0, rect.y, 1.0, rect.h, s);
    }

    pub fn handle_click(&mut self, mx: f32, my: f32, rect: Rect) -> bool {
        if !rect.contains(mx, my) { return false; }

        let items = build_items();
        let list_top = rect.y + HEADER_H;
        let mut y = list_top - self.scroll_y;
        let mut entry_idx = 0usize;

        for item in &items {
            match item {
                ToolItem::Header(_) => { y += SECTION_H; }
                ToolItem::Entry(_) => {
                    if my >= y && my < y + ITEM_H && my >= list_top {
                        self.selected_idx = Some(entry_idx);
                        return true;
                    }
                    entry_idx += 1;
                    y += ITEM_H;
                }
            }
        }
        false
    }
}

fn fill(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pix.fill_rect(r, paint, Transform::identity(), None);
    }
}
