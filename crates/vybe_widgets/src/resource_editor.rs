//! Resource Editor widget — tabbed grid for managing project resources
//! (strings, images, icons, audio, files). Matches Visual Studio's resource editor.

use cosmic_text::{Color as CosmicColor, FontSystem, SwashCache};
use tiny_skia::{Paint, Pixmap, Transform, PathBuilder, Stroke};
use crate::layout::{LayoutRect, RenderContext, MouseEvent, MouseEventKind, KeyEvent, PanelWidget, WidgetEvent, WidgetId, WidgetCommand, CommandValue};
use winit::window::CursorIcon;

const HEADER_H: f32 = 28.0;
const TAB_H: f32 = 26.0;
const ROW_H: f32 = 26.0;
const COL_HEADER_H: f32 = 24.0;
const MIN_COL_W: f32 = 60.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ResourceTab {
    Strings,
    Images,
    Icons,
    Audio,
    Files,
    Other,
}

impl ResourceTab {
    pub fn all() -> &'static [ResourceTab] {
        &[ResourceTab::Strings, ResourceTab::Images, ResourceTab::Icons, ResourceTab::Audio, ResourceTab::Files, ResourceTab::Other]
    }

    pub fn label(&self) -> &'static str {
        match self {
            ResourceTab::Strings => "Strings",
            ResourceTab::Images => "Images",
            ResourceTab::Icons => "Icons",
            ResourceTab::Audio => "Audio",
            ResourceTab::Files => "Files",
            ResourceTab::Other => "Other",
        }
    }

    pub fn file_extensions(&self) -> &'static [&'static str] {
        match self {
            ResourceTab::Images => &["png", "jpg", "jpeg", "gif", "bmp", "tiff", "webp"],
            ResourceTab::Icons => &["ico"],
            ResourceTab::Audio => &["wav", "mp3", "ogg", "flac", "aiff"],
            ResourceTab::Files => &["*"],
            _ => &[],
        }
    }

    pub fn is_file_based(&self) -> bool {
        matches!(self, ResourceTab::Images | ResourceTab::Icons | ResourceTab::Audio | ResourceTab::Files)
    }
}

/// A single resource entry displayed in the grid.
#[derive(Debug, Clone)]
pub struct ResourceEntry {
    pub name: String,
    pub value: String,
    pub comment: String,
    pub tab: ResourceTab,
    /// Original file name for file-based resources
    pub file_name: Option<String>,
}

/// Event returned from click handling.
#[derive(Debug, Clone)]
pub enum ResourceEditorEvent {
    None,
    /// User wants to add a resource — for file-based tabs, IDE should open file picker
    AddResource(ResourceTab),
    /// User wants to delete resource at index
    DeleteResource(usize),
    /// User clicked a cell for editing: (row_index, column: 0=name, 1=value, 2=comment)
    EditCell(usize, usize),
    /// User wants to browse for a file (row_index) — IDE should open single-file picker
    BrowseFile(usize),
    /// Tab changed
    TabChanged(ResourceTab),
    /// Editing committed (row, col, new_value) — IDE should sync to project
    EditCommitted(usize, usize, String),
    /// Add string resource with name/value/comment — IDE should create ResourceItem
    AddStringResource(String, String, String),
}

pub struct ResourceEditor {
    pub id: WidgetId,
    pub entries: Vec<ResourceEntry>,
    pub active_tab: ResourceTab,
    pub scroll_y: f32,
    pub selected_row: Option<usize>,
    /// (row, col, current_text, cursor_pos)
    pub editing: Option<(usize, usize, String, usize)>,
    /// Column width ratios [name, value, comment, actions] — fractions of total width
    pub col_ratios: [f32; 4],
    /// Column being resized: (col_index, start_mx, original_ratios)
    col_resize: Option<(usize, f32, [f32; 4])>,
    /// Add-row fields for string/other tabs: (name, value, comment)
    pub add_fields: (String, String, String),
    /// Which add field is being edited (0=name, 1=value, 2=comment, None=nothing)
    pub add_editing: Option<usize>,
    /// Cursor position within add field
    pub add_cursor: usize,
    /// Whether the resource editor has been modified since last sync
    pub dirty: bool,
    // Layout (PanelWidget)
    pub layout_rect: LayoutRect,
    pub pending_events: Vec<WidgetEvent>,
}

impl ResourceEditor {
    pub fn new() -> Self {
        Self {
            id: WidgetId::next(),
            entries: Vec::new(),
            active_tab: ResourceTab::Strings,
            scroll_y: 0.0,
            selected_row: None,
            editing: None,
            col_ratios: [0.28, 0.38, 0.24, 0.10],
            col_resize: None,
            add_fields: (String::new(), String::new(), String::new()),
            add_editing: None,
            add_cursor: 0,
            dirty: false,
            layout_rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn filtered_indices(&self) -> Vec<usize> {
        self.entries.iter().enumerate()
            .filter(|(_, e)| e.tab == self.active_tab)
            .map(|(i, _)| i)
            .collect()
    }

    fn tab_count(&self, tab: ResourceTab) -> usize {
        self.entries.iter().filter(|e| e.tab == tab).count()
    }

    fn content_h(&self) -> f32 {
        let count = self.filtered_indices().len();
        let add_row = if !self.active_tab.is_file_based() { 1.0 } else { 0.0 };
        COL_HEADER_H + (count as f32 + add_row + 1.0) * ROW_H
    }

    fn col_widths(&self, w: f32) -> (f32, f32, f32, f32) {
        (
            w * self.col_ratios[0],
            w * self.col_ratios[1],
            w * self.col_ratios[2],
            w * self.col_ratios[3],
        )
    }

    /// Get the x positions of column separators (for resize hit testing)
    fn col_separator_xs(&self, x: f32, w: f32) -> [f32; 3] {
        let (name_w, value_w, comment_w, _) = self.col_widths(w);
        [
            x + name_w,
            x + name_w + value_w,
            x + name_w + value_w + comment_w,
        ]
    }

    pub fn scroll(&mut self, delta: f32, visible_h: f32) {
        let max = (self.content_h() - visible_h).max(0.0);
        self.scroll_y = (self.scroll_y - delta).clamp(0.0, max);
    }

    /// Returns display value for a file-based entry (just filename, not full path)
    fn display_value(entry: &ResourceEntry) -> &str {
        if entry.tab.is_file_based() {
            // Show just the filename
            let v = &entry.value;
            v.rsplit(|c| c == '/' || c == '\\').next().unwrap_or(v)
        } else {
            &entry.value
        }
    }

    pub fn render_at(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        x: f32, y: f32, w: f32, h: f32, scale: f32,
    ) {
        let s = scale;
        let mut paint = Paint::default();

        // Background
        paint.set_color_rgba8(255, 255, 255, 255);
        fill(pix, &paint, x, y, w, h, s);

        // Title bar with "Add" button for file-based tabs
        paint.set_color_rgba8(240, 240, 240, 255);
        fill(pix, &paint, x, y, w, HEADER_H, s);
        let title_col = CosmicColor::rgba(50, 50, 50, 255);
        crate::ide_text::draw_text(pix, fs, sc, "Resource Editor", x + 10.0, y + 6.0, 13.0, title_col, s);

        // Add button in title bar for file-based tabs
        if self.active_tab.is_file_based() {
            let btn_label = format!("+ Add {}...", self.active_tab.label());
            let btn_x = w - 140.0;
            paint.set_color_rgba8(0, 120, 212, 255);
            fill(pix, &paint, x + btn_x, y + 3.0, 130.0, HEADER_H - 6.0, s);
            crate::ide_text::draw_text(pix, fs, sc, &btn_label, x + btn_x + 8.0, y + 7.0, 11.0, CosmicColor::rgba(255, 255, 255, 255), s);
        }

        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, x, y + HEADER_H - 1.0, w, 1.0, s);

        // Tabs with counts
        let tab_y = y + HEADER_H;
        let tabs = ResourceTab::all();
        let tab_w = (w / tabs.len() as f32).min(100.0);
        let text_col = CosmicColor::rgba(30, 30, 30, 255);
        let dim_col = CosmicColor::rgba(100, 100, 100, 255);

        for (i, tab) in tabs.iter().enumerate() {
            let tx = x + i as f32 * tab_w;
            let active = *tab == self.active_tab;
            if active {
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, tx, tab_y, tab_w, TAB_H, s);
                paint.set_color_rgba8(0, 120, 212, 255);
                fill(pix, &paint, tx, tab_y + TAB_H - 2.0, tab_w, 2.0, s);
            } else {
                paint.set_color_rgba8(245, 245, 245, 255);
                fill(pix, &paint, tx, tab_y, tab_w, TAB_H, s);
            }
            let count = self.tab_count(*tab);
            let label = if count > 0 {
                format!("{} ({})", tab.label(), count)
            } else {
                tab.label().to_string()
            };
            let col = if active { text_col } else { dim_col };
            crate::ide_text::draw_text(pix, fs, sc, &label, tx + 6.0, tab_y + 5.0, 11.0, col, s);
        }
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, x, tab_y + TAB_H - 1.0, w, 1.0, s);

        // Grid area
        let grid_y = tab_y + TAB_H;
        let grid_h = h - HEADER_H - TAB_H;
        let filtered = self.filtered_indices();
        let (name_w, value_w, comment_w, action_w) = self.col_widths(w);

        // Column headers
        paint.set_color_rgba8(240, 240, 240, 255);
        fill(pix, &paint, x, grid_y, w, COL_HEADER_H, s);
        let header_col = CosmicColor::rgba(60, 60, 60, 255);
        crate::ide_text::draw_text(pix, fs, sc, "Name", x + 6.0, grid_y + 4.0, 11.0, header_col, s);
        let val_header = if self.active_tab.is_file_based() { "File Path" } else { "Value" };
        crate::ide_text::draw_text(pix, fs, sc, val_header, x + name_w + 6.0, grid_y + 4.0, 11.0, header_col, s);
        crate::ide_text::draw_text(pix, fs, sc, "Comment", x + name_w + value_w + 6.0, grid_y + 4.0, 11.0, header_col, s);
        crate::ide_text::draw_text(pix, fs, sc, "Actions", x + name_w + value_w + comment_w + 4.0, grid_y + 4.0, 11.0, header_col, s);

        // Header bottom line
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, x, grid_y + COL_HEADER_H - 1.0, w, 1.0, s);

        // Data rows
        let rows_y = grid_y + COL_HEADER_H;
        let val_col = CosmicColor::rgba(30, 30, 30, 255);
        let del_col = CosmicColor::rgba(200, 60, 60, 255);

        for (row_idx, &entry_idx) in filtered.iter().enumerate() {
            let ry = rows_y + row_idx as f32 * ROW_H - self.scroll_y;
            if ry + ROW_H < rows_y || ry > grid_y + grid_h { continue; }

            let entry = &self.entries[entry_idx];

            // Selection highlight
            if self.selected_row == Some(entry_idx) {
                paint.set_color_rgba8(227, 242, 253, 255);
                fill(pix, &paint, x, ry, w, ROW_H, s);
            } else if row_idx % 2 == 1 {
                paint.set_color_rgba8(250, 250, 250, 255);
                fill(pix, &paint, x, ry, w, ROW_H, s);
            }

            // Check if we're editing this row
            let is_editing = self.editing.as_ref().map(|(r, _, _, _)| *r == entry_idx).unwrap_or(false);

            // Name column
            if is_editing && self.editing.as_ref().unwrap().1 == 0 {
                let (_, _, ref text, cursor) = self.editing.as_ref().unwrap();
                self.render_edit_cell(pix, fs, sc, x, ry, name_w, text, *cursor, s);
            } else {
                crate::ide_text::draw_text(pix, fs, sc, &entry.name, x + 4.0, ry + 5.0, 11.0, val_col, s);
            }

            // Value column
            if is_editing && self.editing.as_ref().unwrap().1 == 1 {
                let (_, _, ref text, cursor) = self.editing.as_ref().unwrap();
                self.render_edit_cell(pix, fs, sc, x + name_w, ry, value_w, text, *cursor, s);
            } else {
                let display = Self::display_value(entry);
                crate::ide_text::draw_text(pix, fs, sc, display, x + name_w + 4.0, ry + 5.0, 11.0, val_col, s);
                // Browse button for file-based resources
                if self.active_tab.is_file_based() {
                    let btn_x = x + name_w + value_w - 56.0;
                    paint.set_color_rgba8(240, 240, 240, 255);
                    fill(pix, &paint, btn_x, ry + 2.0, 52.0, ROW_H - 4.0, s);
                    paint.set_color_rgba8(180, 180, 180, 255);
                    stroke_rect(pix, &paint, btn_x, ry + 2.0, 52.0, ROW_H - 4.0, s);
                    crate::ide_text::draw_text(pix, fs, sc, "Browse", btn_x + 6.0, ry + 5.0, 10.0, CosmicColor::rgba(60, 60, 60, 255), s);
                }
            }

            // Comment column
            if is_editing && self.editing.as_ref().unwrap().1 == 2 {
                let (_, _, ref text, cursor) = self.editing.as_ref().unwrap();
                self.render_edit_cell(pix, fs, sc, x + name_w + value_w, ry, comment_w, text, *cursor, s);
            } else {
                crate::ide_text::draw_text(pix, fs, sc, &entry.comment, x + name_w + value_w + 4.0, ry + 5.0, 11.0, dim_col, s);
            }

            // Delete button
            let del_x = x + name_w + value_w + comment_w + (action_w - 20.0) / 2.0;
            crate::ide_text::draw_text(pix, fs, sc, "\u{2716}", del_x, ry + 5.0, 12.0, del_col, s);

            // Row bottom line
            paint.set_color_rgba8(235, 235, 235, 255);
            fill(pix, &paint, x, ry + ROW_H - 1.0, w, 1.0, s);
        }

        // Add row for string/other tabs — inline fields
        if !self.active_tab.is_file_based() {
            let add_y = rows_y + filtered.len() as f32 * ROW_H - self.scroll_y;
            if add_y + ROW_H > rows_y && add_y < grid_y + grid_h {
                // Separator
                paint.set_color_rgba8(220, 220, 220, 255);
                fill(pix, &paint, x, add_y - 1.0, w, 1.0, s);

                // Background
                paint.set_color_rgba8(252, 252, 252, 255);
                fill(pix, &paint, x, add_y, w, ROW_H, s);

                let placeholder_col = CosmicColor::rgba(160, 160, 160, 255);

                // Name field
                if self.add_editing == Some(0) {
                    self.render_edit_cell(pix, fs, sc, x, add_y, name_w, &self.add_fields.0, self.add_cursor, s);
                } else if self.add_fields.0.is_empty() {
                    crate::ide_text::draw_text(pix, fs, sc, "Name", x + 4.0, add_y + 5.0, 11.0, placeholder_col, s);
                } else {
                    crate::ide_text::draw_text(pix, fs, sc, &self.add_fields.0, x + 4.0, add_y + 5.0, 11.0, val_col, s);
                }

                // Value field
                if self.add_editing == Some(1) {
                    self.render_edit_cell(pix, fs, sc, x + name_w, add_y, value_w, &self.add_fields.1, self.add_cursor, s);
                } else if self.add_fields.1.is_empty() {
                    crate::ide_text::draw_text(pix, fs, sc, "Value", x + name_w + 4.0, add_y + 5.0, 11.0, placeholder_col, s);
                } else {
                    crate::ide_text::draw_text(pix, fs, sc, &self.add_fields.1, x + name_w + 4.0, add_y + 5.0, 11.0, val_col, s);
                }

                // Comment field
                if self.add_editing == Some(2) {
                    self.render_edit_cell(pix, fs, sc, x + name_w + value_w, add_y, comment_w, &self.add_fields.2, self.add_cursor, s);
                } else if self.add_fields.2.is_empty() {
                    crate::ide_text::draw_text(pix, fs, sc, "Comment", x + name_w + value_w + 4.0, add_y + 5.0, 11.0, placeholder_col, s);
                } else {
                    crate::ide_text::draw_text(pix, fs, sc, &self.add_fields.2, x + name_w + value_w + 4.0, add_y + 5.0, 11.0, dim_col, s);
                }

                // Add button in actions column
                let add_btn_x = x + name_w + value_w + comment_w + 4.0;
                paint.set_color_rgba8(0, 120, 212, 255);
                fill(pix, &paint, add_btn_x, add_y + 3.0, action_w - 8.0, ROW_H - 6.0, s);
                crate::ide_text::draw_text(pix, fs, sc, "+ Add", add_btn_x + 4.0, add_y + 6.0, 10.0, CosmicColor::rgba(255, 255, 255, 255), s);
            }
        }

        // Column separator lines — drawn AFTER rows so they stay on top
        paint.set_color_rgba8(200, 200, 200, 255);
        fill(pix, &paint, x + name_w - 1.0, grid_y, 2.0, grid_h, s);
        fill(pix, &paint, x + name_w + value_w - 1.0, grid_y, 2.0, grid_h, s);
        fill(pix, &paint, x + name_w + value_w + comment_w - 1.0, grid_y, 2.0, grid_h, s);

        // Overdraw header/tabs area to clip scrolled rows that bleed above grid
        paint.set_color_rgba8(240, 240, 240, 255);
        fill(pix, &paint, x, y, w, HEADER_H, s);
        crate::ide_text::draw_text(pix, fs, sc, "Resource Editor", x + 10.0, y + 6.0, 13.0, title_col, s);
        if self.active_tab.is_file_based() {
            let btn_label = format!("+ Add {}...", self.active_tab.label());
            let btn_x = w - 140.0;
            paint.set_color_rgba8(0, 120, 212, 255);
            fill(pix, &paint, x + btn_x, y + 3.0, 130.0, HEADER_H - 6.0, s);
            crate::ide_text::draw_text(pix, fs, sc, &btn_label, x + btn_x + 8.0, y + 7.0, 11.0, CosmicColor::rgba(255, 255, 255, 255), s);
        }
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, x, y + HEADER_H - 1.0, w, 1.0, s);

        // Re-render tabs on top (clip fix)
        for (i, tab) in tabs.iter().enumerate() {
            let tx = x + i as f32 * tab_w;
            let active = *tab == self.active_tab;
            if active {
                paint.set_color_rgba8(255, 255, 255, 255);
            } else {
                paint.set_color_rgba8(245, 245, 245, 255);
            }
            fill(pix, &paint, tx, tab_y, tab_w, TAB_H, s);
            if active {
                paint.set_color_rgba8(0, 120, 212, 255);
                fill(pix, &paint, tx, tab_y + TAB_H - 2.0, tab_w, 2.0, s);
            }
            let count = self.tab_count(*tab);
            let label = if count > 0 {
                format!("{} ({})", tab.label(), count)
            } else {
                tab.label().to_string()
            };
            let col = if active { text_col } else { dim_col };
            crate::ide_text::draw_text(pix, fs, sc, &label, tx + 6.0, tab_y + 5.0, 11.0, col, s);
        }
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, x, tab_y + TAB_H - 1.0, w, 1.0, s);

        // Outer border
        paint.set_color_rgba8(204, 204, 204, 255);
        stroke_rect(pix, &paint, x, y, w, h, s);
    }

    fn render_edit_cell(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        cx: f32, cy: f32, cw: f32, text: &str, cursor: usize, scale: f32,
    ) {
        let mut paint = Paint::default();
        let val_col = CosmicColor::rgba(30, 30, 30, 255);

        // White background
        paint.set_color_rgba8(255, 255, 255, 255);
        fill(pix, &paint, cx + 1.0, cy + 1.0, cw - 2.0, ROW_H - 2.0, scale);
        // Blue border
        paint.set_color_rgba8(0, 120, 212, 255);
        stroke_rect(pix, &paint, cx + 1.0, cy + 1.0, cw - 2.0, ROW_H - 2.0, scale);
        // Text
        crate::ide_text::draw_text(pix, fs, sc, text, cx + 4.0, cy + 5.0, 11.0, val_col, scale);
        // Cursor — measure text before cursor for accurate positioning
        let before_cursor = if cursor <= text.len() { &text[..cursor] } else { text };
        let cursor_offset = crate::ide_text::measure_text(fs, before_cursor, 11.0, scale);
        let cursor_x = cx + 4.0 + cursor_offset;
        paint.set_color_rgba8(0, 0, 0, 255);
        fill(pix, &paint, cursor_x, cy + 3.0, 1.5, ROW_H - 6.0, scale);
    }

    pub fn handle_click(&mut self, mx: f32, my: f32, x: f32, y: f32, w: f32, _h: f32) -> ResourceEditorEvent {
        let (name_w, value_w, comment_w, _action_w) = self.col_widths(w);

        // Add button in title bar for file-based tabs
        if self.active_tab.is_file_based() && my >= y && my < y + HEADER_H {
            let btn_x = x + w - 140.0;
            if mx >= btn_x && mx < btn_x + 130.0 {
                return ResourceEditorEvent::AddResource(self.active_tab);
            }
        }

        // Tab click
        let tab_y = y + HEADER_H;
        if my >= tab_y && my < tab_y + TAB_H {
            let tabs = ResourceTab::all();
            let tab_w = (w / tabs.len() as f32).min(100.0);
            let idx = ((mx - x) / tab_w) as usize;
            if idx < tabs.len() {
                // Commit any pending edit first
                self.commit_edit();
                self.active_tab = tabs[idx];
                self.scroll_y = 0.0;
                self.selected_row = None;
                self.editing = None;
                self.add_editing = None;
                return ResourceEditorEvent::TabChanged(self.active_tab);
            }
        }

        let grid_y = tab_y + TAB_H;
        let rows_y = grid_y + COL_HEADER_H;
        let filtered = self.filtered_indices();

        // Check data rows
        for (row_idx, &entry_idx) in filtered.iter().enumerate() {
            let ry = rows_y + row_idx as f32 * ROW_H - self.scroll_y;
            if my >= ry && my < ry + ROW_H {
                self.selected_row = Some(entry_idx);
                self.add_editing = None;

                // Delete button
                let del_x = x + name_w + value_w + comment_w;
                if mx >= del_x {
                    self.commit_edit();
                    self.editing = None;
                    return ResourceEditorEvent::DeleteResource(entry_idx);
                }

                // Browse button for file-based resources
                if self.active_tab.is_file_based() && mx >= x + name_w + value_w - 56.0 && mx < x + name_w + value_w {
                    self.commit_edit();
                    return ResourceEditorEvent::BrowseFile(entry_idx);
                }

                // Cell click for editing
                let col = if mx < x + name_w { 0 }
                    else if mx < x + name_w + value_w { 1 }
                    else if mx < x + name_w + value_w + comment_w { 2 }
                    else { 3 };

                if col < 3 {
                    // For file-based tabs, value column is not directly editable (use Browse)
                    if col == 1 && self.active_tab.is_file_based() {
                        return ResourceEditorEvent::None;
                    }
                    // Commit previous edit if switching cells
                    self.commit_edit();
                    let text = match col {
                        0 => self.entries[entry_idx].name.clone(),
                        1 => self.entries[entry_idx].value.clone(),
                        2 => self.entries[entry_idx].comment.clone(),
                        _ => String::new(),
                    };
                    let cursor = text.len();
                    self.editing = Some((entry_idx, col, text, cursor));
                    return ResourceEditorEvent::EditCell(entry_idx, col);
                }

                return ResourceEditorEvent::None;
            }
        }

        // Add row click (string/other tabs)
        if !self.active_tab.is_file_based() {
            let add_y = rows_y + filtered.len() as f32 * ROW_H - self.scroll_y;
            if my >= add_y && my < add_y + ROW_H {
                // Which column?
                let col = if mx < x + name_w { 0 }
                    else if mx < x + name_w + value_w { 1 }
                    else if mx < x + name_w + value_w + comment_w { 2 }
                    else { 3 };

                if col == 3 {
                    // Add button clicked
                    if !self.add_fields.0.is_empty() {
                        let evt = ResourceEditorEvent::AddStringResource(
                            self.add_fields.0.clone(),
                            self.add_fields.1.clone(),
                            self.add_fields.2.clone(),
                        );
                        self.add_fields = (String::new(), String::new(), String::new());
                        self.add_editing = None;
                        return evt;
                    }
                } else {
                    // Edit add field
                    self.commit_edit();
                    self.editing = None;
                    self.selected_row = None;
                    let text = match col {
                        0 => &self.add_fields.0,
                        1 => &self.add_fields.1,
                        2 => &self.add_fields.2,
                        _ => &self.add_fields.0,
                    };
                    self.add_cursor = text.len();
                    self.add_editing = Some(col);
                }
                return ResourceEditorEvent::None;
            }
        }

        // Click outside rows — deselect and commit
        self.commit_edit();
        self.selected_row = None;
        self.editing = None;
        self.add_editing = None;
        ResourceEditorEvent::None
    }

    /// Handle column resize drag start — returns true if we started a resize
    pub fn handle_col_resize_start(&mut self, mx: f32, my: f32, x: f32, y: f32, w: f32, h: f32) -> bool {
        let tab_y = y + HEADER_H;
        let grid_y = tab_y + TAB_H;
        // Allow resize anywhere in the grid area (header + data rows)
        if my < grid_y || my > y + h { return false; }
        let seps = self.col_separator_xs(x, w);
        for (i, &sep_x) in seps.iter().enumerate() {
            if (mx - sep_x).abs() < 6.0 {
                self.col_resize = Some((i, mx, self.col_ratios));
                return true;
            }
        }
        false
    }

    /// Handle column resize drag
    pub fn handle_col_resize_move(&mut self, mx: f32, w: f32) {
        if let Some((col_idx, start_mx, orig_ratios)) = self.col_resize {
            let delta = (mx - start_mx) / w;
            let mut new_ratios = orig_ratios;
            // Adjust the dragged column and next column
            let min_ratio = MIN_COL_W / w;
            new_ratios[col_idx] = (orig_ratios[col_idx] + delta).max(min_ratio);
            new_ratios[col_idx + 1] = (orig_ratios[col_idx + 1] - delta).max(min_ratio);
            // Normalize so they sum to 1.0
            let sum: f32 = new_ratios.iter().sum();
            if sum > 0.0 {
                for r in &mut new_ratios {
                    *r /= sum;
                }
            }
            self.col_ratios = new_ratios;
        }
    }

    /// End column resize
    pub fn handle_col_resize_end(&mut self) {
        self.col_resize = None;
    }

    pub fn is_resizing(&self) -> bool {
        self.col_resize.is_some()
    }

    /// Check if mouse is near a column separator (for cursor change)
    pub fn is_near_separator(&self, mx: f32, x: f32, y: f32, w: f32, my: f32, h: f32) -> bool {
        let tab_y = y + HEADER_H;
        let grid_y = tab_y + TAB_H;
        if my < grid_y || my > y + h { return false; }
        let seps = self.col_separator_xs(x, w);
        seps.iter().any(|&sep_x| (mx - sep_x).abs() < 6.0)
    }

    /// Handle character input for active editing cell or add-row field
    pub fn handle_key(&mut self, ch: char) {
        if let Some((_, _, ref mut text, ref mut cursor)) = self.editing {
            if ch == '\x08' { // backspace
                if *cursor > 0 {
                    text.remove(*cursor - 1);
                    *cursor -= 1;
                }
            } else if !ch.is_control() {
                text.insert(*cursor, ch);
                *cursor += 1;
            }
        } else if let Some(col) = self.add_editing {
            let text = match col {
                0 => &mut self.add_fields.0,
                1 => &mut self.add_fields.1,
                2 => &mut self.add_fields.2,
                _ => return,
            };
            if ch == '\x08' {
                if self.add_cursor > 0 {
                    text.remove(self.add_cursor - 1);
                    self.add_cursor -= 1;
                }
            } else if !ch.is_control() {
                text.insert(self.add_cursor, ch);
                self.add_cursor += 1;
            }
        }
    }

    /// Handle Enter key — commit edit or add resource
    pub fn handle_enter(&mut self) -> ResourceEditorEvent {
        if self.editing.is_some() {
            return self.commit_edit_with_event();
        }
        if self.add_editing.is_some() && !self.add_fields.0.is_empty() {
            let evt = ResourceEditorEvent::AddStringResource(
                self.add_fields.0.clone(),
                self.add_fields.1.clone(),
                self.add_fields.2.clone(),
            );
            self.add_fields = (String::new(), String::new(), String::new());
            self.add_editing = None;
            return evt;
        }
        ResourceEditorEvent::None
    }

    /// Handle Escape key — cancel editing
    pub fn handle_escape(&mut self) {
        self.editing = None;
        self.add_editing = None;
    }

    /// Handle Tab key — move to next cell
    pub fn handle_tab(&mut self) -> ResourceEditorEvent {
        if let Some((row, col, _, _)) = self.editing {
            let evt = self.commit_edit_with_event();
            let next_col = col + 1;
            if next_col < 3 {
                // Move to next column in same row
                // Skip value column for file-based tabs
                let actual_col = if next_col == 1 && self.active_tab.is_file_based() { 2 } else { next_col };
                if actual_col < 3 && row < self.entries.len() {
                    let text = match actual_col {
                        0 => self.entries[row].name.clone(),
                        1 => self.entries[row].value.clone(),
                        2 => self.entries[row].comment.clone(),
                        _ => String::new(),
                    };
                    let cursor = text.len();
                    self.editing = Some((row, actual_col, text, cursor));
                }
            } else {
                self.editing = None;
            }
            return evt;
        }
        if let Some(col) = self.add_editing {
            let next = col + 1;
            if next < 3 {
                let text = match next {
                    0 => &self.add_fields.0,
                    1 => &self.add_fields.1,
                    2 => &self.add_fields.2,
                    _ => &self.add_fields.0,
                };
                self.add_cursor = text.len();
                self.add_editing = Some(next);
            } else {
                self.add_editing = None;
            }
        }
        ResourceEditorEvent::None
    }

    /// Handle Delete key — delete selected row
    pub fn handle_delete(&mut self) -> ResourceEditorEvent {
        if self.editing.is_some() || self.add_editing.is_some() {
            return ResourceEditorEvent::None; // Don't delete while editing
        }
        if let Some(idx) = self.selected_row {
            if idx < self.entries.len() {
                self.selected_row = None;
                return ResourceEditorEvent::DeleteResource(idx);
            }
        }
        ResourceEditorEvent::None
    }

    /// Handle Home key in editing
    pub fn handle_home(&mut self) {
        if let Some((_, _, _, ref mut cursor)) = self.editing {
            *cursor = 0;
        } else if self.add_editing.is_some() {
            self.add_cursor = 0;
        }
    }

    /// Handle End key in editing
    pub fn handle_end(&mut self) {
        if let Some((_, _, ref text, ref mut cursor)) = self.editing {
            *cursor = text.len();
        } else if let Some(col) = self.add_editing {
            let text = match col {
                0 => &self.add_fields.0,
                1 => &self.add_fields.1,
                2 => &self.add_fields.2,
                _ => return,
            };
            self.add_cursor = text.len();
        }
    }

    /// Handle Left arrow
    pub fn handle_left(&mut self) {
        if let Some((_, _, _, ref mut cursor)) = self.editing {
            if *cursor > 0 { *cursor -= 1; }
        } else if self.add_editing.is_some() {
            if self.add_cursor > 0 { self.add_cursor -= 1; }
        }
    }

    /// Handle Right arrow
    pub fn handle_right(&mut self) {
        if let Some((_, _, ref text, ref mut cursor)) = self.editing {
            if *cursor < text.len() { *cursor += 1; }
        } else if let Some(col) = self.add_editing {
            let len = match col {
                0 => self.add_fields.0.len(),
                1 => self.add_fields.1.len(),
                2 => self.add_fields.2.len(),
                _ => 0,
            };
            if self.add_cursor < len { self.add_cursor += 1; }
        }
    }

    /// Commit the current edit and update the entry
    pub fn commit_edit(&mut self) {
        if let Some((row, col, ref text, _)) = self.editing {
            if row < self.entries.len() {
                match col {
                    0 => self.entries[row].name = text.clone(),
                    1 => self.entries[row].value = text.clone(),
                    2 => self.entries[row].comment = text.clone(),
                    _ => {}
                }
                self.dirty = true;
            }
        }
        self.editing = None;
    }

    /// Commit edit and return event for IDE to sync to project
    fn commit_edit_with_event(&mut self) -> ResourceEditorEvent {
        if let Some((row, col, ref text, _)) = self.editing.clone() {
            if row < self.entries.len() {
                match col {
                    0 => self.entries[row].name = text.clone(),
                    1 => self.entries[row].value = text.clone(),
                    2 => self.entries[row].comment = text.clone(),
                    _ => {}
                }
                self.dirty = true;
                self.editing = None;
                return ResourceEditorEvent::EditCommitted(row, col, text.clone());
            }
        }
        self.editing = None;
        ResourceEditorEvent::None
    }

    /// Whether any field is being edited (cell or add row)
    pub fn is_editing(&self) -> bool {
        self.editing.is_some() || self.add_editing.is_some()
    }
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
        let mut st = Stroke::default(); st.width = 1.0 * s;
        pix.stroke_path(&path, paint, &st, Transform::identity(), None);
    }
}

// ── PanelWidget impl ───────────────────────────────────────────────────

impl PanelWidget for ResourceEditor {
    fn set_rect(&mut self, rect: LayoutRect) { self.layout_rect = rect; }
    fn rect(&self) -> LayoutRect { self.layout_rect }
    fn widget_id(&self) -> WidgetId { self.id }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.layout_rect;
        self.render_at(ctx.pixmap, ctx.font_system, ctx.swash_cache, r.x, r.y, r.w, r.h, ctx.scale);
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.layout_rect.contains(event.x, event.y) { return false; }
        true
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        if self.layout_rect.contains(x, y) {
            let visible_h = self.layout_rect.h;
            self.scroll(delta, visible_h);
            true
        } else {
            false
        }
    }

    fn cursor_at(&self, x: f32, y: f32) -> CursorIcon {
        if self.layout_rect.contains(x, y) {
            let r = self.layout_rect;
            if self.is_resizing() || self.is_near_separator(x, r.x, r.y, r.w, y, r.h) {
                CursorIcon::ColResize
            } else {
                CursorIcon::Default
            }
        } else {
            CursorIcon::Default
        }
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }

    fn focusable(&self) -> bool { true }
}
