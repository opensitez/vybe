//! Project Explorer panel — shows project structure (forms, code files, references).

use cosmic_text::{Color as CosmicColor, FontSystem, SwashCache};
use tiny_skia::{Paint, Pixmap, Transform};
use vybe_project::project::Project;

use crate::layout::Rect;
use crate::text::draw_text;

pub enum ExplorerEvent {
    SelectForm(String),
    SelectCode(String),
    ViewCode(String),
    None,
}

const HEADER_H: f32 = 28.0;
const ITEM_H: f32 = 24.0;
const INDENT: f32 = 16.0;
const SCROLLBAR_W: f32 = 10.0;

pub struct ProjectExplorer {
    pub scroll_y: f32,
    pub forms_collapsed: bool,
    pub code_collapsed: bool,
    pub refs_collapsed: bool,
}

impl ProjectExplorer {
    pub fn new() -> Self {
        Self {
            scroll_y: 0.0,
            forms_collapsed: false,
            code_collapsed: false,
            refs_collapsed: false,
        }
    }

    fn content_h(&self, project: &Project) -> f32 {
        let mut n = 1; // project name row
        // Forms section
        n += 1; // "Forms" header
        if !self.forms_collapsed {
            n += project.forms.len();
        }
        // Code section
        if !project.code_files.is_empty() {
            n += 1; // "Code" header
            if !self.code_collapsed {
                n += project.code_files.len();
            }
        }
        // References section
        if !project.project_references.is_empty() {
            n += 1;
            if !self.refs_collapsed {
                n += project.project_references.len();
            }
        }
        n as f32 * ITEM_H
    }

    fn max_scroll(&self, rect: Rect, project: &Project) -> f32 {
        (self.content_h(project) - (rect.h - HEADER_H)).max(0.0)
    }

    pub fn scroll(&mut self, delta: f32, rect: Rect, project: &Project) {
        self.scroll_y = (self.scroll_y - delta * ITEM_H * 3.0)
            .max(0.0)
            .min(self.max_scroll(rect, project));
    }

    pub fn render(
        &self,
        pix: &mut Pixmap,
        fs: &mut FontSystem,
        sc: &mut SwashCache,
        rect: Rect,
        scale: f32,
        project: &Project,
        current_form: Option<&str>,
    ) {
        let s = scale;
        let mut paint = Paint::default();

        // Background
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);

        // Title
        let title_color = CosmicColor::rgba(50, 50, 50, 255);
        draw_text(pix, fs, sc, "Project Explorer", rect.x + 10.0, rect.y + 6.0, 13.0, title_color, s);

        // Separator
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + HEADER_H - 1.0, rect.w, 1.0, s);

        let text_color = CosmicColor::rgba(30, 30, 30, 255);
        let bold_color = CosmicColor::rgba(30, 30, 30, 255);
        let dim_color = CosmicColor::rgba(100, 100, 100, 255);
        let sel_color = (227u8, 242, 253, 255);
        let list_top = rect.y + HEADER_H;
        let list_h = rect.h - HEADER_H;
        let mut y = list_top - self.scroll_y;

        // Project name
        self.draw_item(pix, fs, sc, &mut paint, &format!("\u{1F4C1} {}", project.name),
            rect.x + 8.0, y, rect.w, bold_color, None, false, list_top, list_top + list_h, s);
        y += ITEM_H;

        // Forms section
        let forms_arrow = if self.forms_collapsed { "\u{25B6}" } else { "\u{25BC}" };
        self.draw_item(pix, fs, sc, &mut paint, &format!("{} \u{1F4CB} Forms", forms_arrow),
            rect.x + 8.0 + INDENT, y, rect.w, bold_color, None, false, list_top, list_top + list_h, s);
        y += ITEM_H;

        if !self.forms_collapsed {
            for fm in &project.forms {
                let is_sel = current_form == Some(fm.form.name.as_str());
                let bg = if is_sel { Some(sel_color) } else { None };
                self.draw_item(pix, fs, sc, &mut paint, &format!("  {}", fm.form.name),
                    rect.x + 8.0 + INDENT * 2.0, y, rect.w, text_color, bg, false, list_top, list_top + list_h, s);
                y += ITEM_H;
            }
        }

        // Code section
        if !project.code_files.is_empty() {
            let code_arrow = if self.code_collapsed { "\u{25B6}" } else { "\u{25BC}" };
            self.draw_item(pix, fs, sc, &mut paint, &format!("{} \u{1F4C4} Code", code_arrow),
                rect.x + 8.0 + INDENT, y, rect.w, bold_color, None, false, list_top, list_top + list_h, s);
            y += ITEM_H;

            if !self.code_collapsed {
                for cf in &project.code_files {
                    let is_sel = current_form == Some(cf.name.as_str());
                    let bg = if is_sel { Some(sel_color) } else { None };
                    self.draw_item(pix, fs, sc, &mut paint, &format!("  {}", cf.name),
                        rect.x + 8.0 + INDENT * 2.0, y, rect.w, text_color, bg, false, list_top, list_top + list_h, s);
                    y += ITEM_H;
                }
            }
        }

        // References section
        if !project.project_references.is_empty() {
            let refs_arrow = if self.refs_collapsed { "\u{25B6}" } else { "\u{25BC}" };
            self.draw_item(pix, fs, sc, &mut paint, &format!("{} \u{1F517} References", refs_arrow),
                rect.x + 8.0 + INDENT, y, rect.w, bold_color, None, false, list_top, list_top + list_h, s);
            y += ITEM_H;

            if !self.refs_collapsed {
                for rn in &project.project_references {
                    self.draw_item(pix, fs, sc, &mut paint, &format!("  {}", rn),
                        rect.x + 8.0 + INDENT * 2.0, y, rect.w, dim_color, None, false, list_top, list_top + list_h, s);
                    y += ITEM_H;
                }
            }
        }

        // Overdraw header to clip scrolled content
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, HEADER_H, s);
        draw_text(pix, fs, sc, "Project Explorer", rect.x + 10.0, rect.y + 6.0, 13.0, title_color, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + HEADER_H - 1.0, rect.w, 1.0, s);

        // Scrollbar
        let max_scroll = self.max_scroll(rect, project);
        if max_scroll > 0.0 {
            let sb_x = rect.x + rect.w - SCROLLBAR_W;
            paint.set_color_rgba8(235, 235, 235, 255);
            fill(pix, &paint, sb_x, list_top, SCROLLBAR_W, list_h, s);
            let content_h = self.content_h(project);
            let visible_frac = (list_h / content_h).min(1.0);
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

    fn draw_item(
        &self,
        pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        paint: &mut Paint, text: &str, x: f32, y: f32, _w: f32,
        color: CosmicColor, bg: Option<(u8, u8, u8, u8)>,
        _bold: bool, clip_top: f32, clip_bottom: f32, s: f32,
    ) {
        if y + ITEM_H < clip_top || y > clip_bottom { return; }
        if let Some((r, g, b, a)) = bg {
            paint.set_color_rgba8(r, g, b, a);
            fill(pix, paint, x - 8.0, y, 300.0, ITEM_H, s);
        }
        draw_text(pix, fs, sc, text, x, y + 4.0, 12.0, color, s);
    }

    pub fn handle_click(&mut self, mx: f32, my: f32, rect: Rect, project: &Project) -> ExplorerEvent {
        if !rect.contains(mx, my) { return ExplorerEvent::None; }

        let list_top = rect.y + HEADER_H;
        let mut y = list_top - self.scroll_y;

        // Project name — skip
        y += ITEM_H;

        // Forms header
        if my >= y && my < y + ITEM_H {
            self.forms_collapsed = !self.forms_collapsed;
            return ExplorerEvent::None;
        }
        y += ITEM_H;

        if !self.forms_collapsed {
            for fm in &project.forms {
                if my >= y && my < y + ITEM_H {
                    return ExplorerEvent::SelectForm(fm.form.name.clone());
                }
                y += ITEM_H;
            }
        }

        // Code header
        if !project.code_files.is_empty() {
            if my >= y && my < y + ITEM_H {
                self.code_collapsed = !self.code_collapsed;
                return ExplorerEvent::None;
            }
            y += ITEM_H;

            if !self.code_collapsed {
                for cf in &project.code_files {
                    if my >= y && my < y + ITEM_H {
                        return ExplorerEvent::SelectCode(cf.name.clone());
                    }
                    y += ITEM_H;
                }
            }
        }

        // References header
        if !project.project_references.is_empty() {
            if my >= y && my < y + ITEM_H {
                self.refs_collapsed = !self.refs_collapsed;
                return ExplorerEvent::None;
            }
            // Ref items are not clickable
        }

        ExplorerEvent::None
    }
}

fn fill(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pix.fill_rect(r, paint, Transform::identity(), None);
    }
}
