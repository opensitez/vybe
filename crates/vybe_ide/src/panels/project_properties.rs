//! Project Properties modal dialog.

use cosmic_text::{Color as CosmicColor, FontSystem, SwashCache};
use tiny_skia::{Paint, Pixmap, Transform};
use vybe_project::project::{Project, StartupObject};

use crate::text::draw_text;

const DIALOG_W: f32 = 400.0;
const DIALOG_H: f32 = 280.0;
const TITLE_H: f32 = 30.0;
const ROW_H: f32 = 28.0;
const FOOTER_H: f32 = 44.0;

pub struct ProjectPropertiesDialog {
    pub visible: bool,
    /// Index into startup options: 0 = Sub Main, 1 = (None), 2.. = form names
    pub selected_startup: usize,
    pub editing_name: bool,
    pub name_value: String,
}

impl ProjectPropertiesDialog {
    pub fn new() -> Self {
        Self {
            visible: false,
            selected_startup: 0,
            editing_name: false,
            name_value: String::new(),
        }
    }

    pub fn open(&mut self, project: &Project) {
        self.visible = true;
        self.name_value = project.name.clone();
        self.selected_startup = match &project.startup_object {
            StartupObject::SubMain => 0,
            StartupObject::None => 1,
            StartupObject::Form(name) => {
                project.forms.iter().position(|f| &f.form.name == name)
                    .map(|i| i + 2)
                    .unwrap_or(0)
            }
        };
    }

    pub fn close(&mut self) {
        self.visible = false;
    }

    /// Apply changes to the project.
    pub fn apply(&self, project: &mut Project) {
        match self.selected_startup {
            0 => {
                project.startup_object = StartupObject::SubMain;
                project.startup_form = None;
            }
            1 => {
                project.startup_object = StartupObject::None;
                project.startup_form = None;
            }
            n => {
                let idx = n - 2;
                if let Some(fm) = project.forms.get(idx) {
                    let name = fm.form.name.clone();
                    project.startup_object = StartupObject::Form(name.clone());
                    project.startup_form = Some(name);
                }
            }
        }
    }

    pub fn render(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        win_w: f32, win_h: f32, scale: f32, project: &Project,
    ) {
        if !self.visible { return; }
        let s = scale;
        let mut paint = Paint::default();

        // Overlay background
        paint.set_color_rgba8(0, 0, 0, 120);
        fill(pix, &paint, 0.0, 0.0, win_w, win_h, s);

        // Dialog position (centered)
        let dx = (win_w - DIALOG_W) / 2.0;
        let dy = (win_h - DIALOG_H) / 2.0;

        // Shadow
        paint.set_color_rgba8(0, 0, 0, 40);
        fill(pix, &paint, dx + 4.0, dy + 4.0, DIALOG_W, DIALOG_H, s);

        // Dialog background
        paint.set_color_rgba8(255, 255, 255, 255);
        fill(pix, &paint, dx, dy, DIALOG_W, DIALOG_H, s);

        // Title bar (blue gradient)
        paint.set_color_rgba8(0, 120, 212, 255);
        fill(pix, &paint, dx, dy, DIALOG_W, TITLE_H, s);

        let white = CosmicColor::rgba(255, 255, 255, 255);
        draw_text(pix, fs, sc, &format!("{} - Project Properties", project.name),
            dx + 10.0, dy + 7.0, 13.0, white, s);

        // Close button
        draw_text(pix, fs, sc, "X", dx + DIALOG_W - 22.0, dy + 7.0, 13.0, white, s);

        // Content
        let text_color = CosmicColor::rgba(30, 30, 30, 255);
        let label_color = CosmicColor::rgba(60, 60, 60, 255);
        let dim_color = CosmicColor::rgba(120, 120, 120, 255);
        let mut y = dy + TITLE_H + 16.0;

        // Project Name
        draw_text(pix, fs, sc, "Project Name:", dx + 16.0, y, 12.0, label_color, s);
        y += 20.0;
        // Name field (read-only style)
        paint.set_color_rgba8(245, 245, 245, 255);
        fill(pix, &paint, dx + 16.0, y, DIALOG_W - 32.0, ROW_H, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        stroke_rect(pix, &paint, dx + 16.0, y, DIALOG_W - 32.0, ROW_H, s);
        draw_text(pix, fs, sc, &project.name, dx + 22.0, y + 6.0, 12.0, text_color, s);
        y += ROW_H + 16.0;

        // Startup Object
        draw_text(pix, fs, sc, "Startup Object:", dx + 16.0, y, 12.0, label_color, s);
        y += 20.0;

        // Startup options as a list
        let options = self.startup_options(project);
        let opt_h = 22.0;
        paint.set_color_rgba8(255, 255, 255, 255);
        let list_h = options.len() as f32 * opt_h;
        fill(pix, &paint, dx + 16.0, y, DIALOG_W - 32.0, list_h, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        stroke_rect(pix, &paint, dx + 16.0, y, DIALOG_W - 32.0, list_h, s);

        for (i, (label, _)) in options.iter().enumerate() {
            let oy = y + i as f32 * opt_h;
            if i == self.selected_startup {
                paint.set_color_rgba8(0, 120, 212, 255);
                fill(pix, &paint, dx + 17.0, oy + 1.0, DIALOG_W - 34.0, opt_h - 2.0, s);
                draw_text(pix, fs, sc, label, dx + 24.0, oy + 4.0, 11.0, white, s);
            } else {
                draw_text(pix, fs, sc, label, dx + 24.0, oy + 4.0, 11.0, text_color, s);
            }
        }

        // Footer
        let footer_y = dy + DIALOG_H - FOOTER_H;
        paint.set_color_rgba8(240, 240, 240, 255);
        fill(pix, &paint, dx, footer_y, DIALOG_W, FOOTER_H, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, dx, footer_y, DIALOG_W, 1.0, s);

        // OK button
        let btn_w = 80.0;
        let btn_h = 28.0;
        let ok_x = dx + DIALOG_W - btn_w * 2.0 - 24.0;
        let btn_y = footer_y + (FOOTER_H - btn_h) / 2.0;
        paint.set_color_rgba8(0, 120, 212, 255);
        fill(pix, &paint, ok_x, btn_y, btn_w, btn_h, s);
        draw_text(pix, fs, sc, "OK", ok_x + 30.0, btn_y + 6.0, 12.0, white, s);

        // Cancel button
        let cancel_x = dx + DIALOG_W - btn_w - 16.0;
        paint.set_color_rgba8(240, 240, 240, 255);
        fill(pix, &paint, cancel_x, btn_y, btn_w, btn_h, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        stroke_rect(pix, &paint, cancel_x, btn_y, btn_w, btn_h, s);
        draw_text(pix, fs, sc, "Cancel", cancel_x + 20.0, btn_y + 6.0, 12.0, text_color, s);
    }

    fn startup_options<'a>(&self, project: &'a Project) -> Vec<(String, usize)> {
        let mut opts = vec![
            ("Sub Main".into(), 0),
            ("(None)".into(), 1),
        ];
        for (i, fm) in project.forms.iter().enumerate() {
            opts.push((fm.form.name.clone(), i + 2));
        }
        opts
    }

    /// Handle click. Returns true if the dialog consumed the event.
    pub fn handle_click(&mut self, mx: f32, my: f32, win_w: f32, win_h: f32, project: &Project) -> bool {
        if !self.visible { return false; }

        let dx = (win_w - DIALOG_W) / 2.0;
        let dy = (win_h - DIALOG_H) / 2.0;

        // Outside dialog = close
        if mx < dx || mx > dx + DIALOG_W || my < dy || my > dy + DIALOG_H {
            self.close();
            return true;
        }

        // Close button
        if mx >= dx + DIALOG_W - 28.0 && mx < dx + DIALOG_W && my >= dy && my < dy + TITLE_H {
            self.close();
            return true;
        }

        // Startup option list
        let list_y = dy + TITLE_H + 16.0 + 20.0 + 28.0 + 16.0 + 20.0;
        let options = self.startup_options(project);
        let opt_h = 22.0;
        if mx >= dx + 16.0 && mx < dx + DIALOG_W - 16.0 {
            let rel_y = my - list_y;
            if rel_y >= 0.0 {
                let idx = (rel_y / opt_h) as usize;
                if idx < options.len() {
                    self.selected_startup = idx;
                    return true;
                }
            }
        }

        // OK button
        let footer_y = dy + DIALOG_H - FOOTER_H;
        let btn_w = 80.0;
        let btn_h = 28.0;
        let ok_x = dx + DIALOG_W - btn_w * 2.0 - 24.0;
        let btn_y = footer_y + (FOOTER_H - btn_h) / 2.0;
        if mx >= ok_x && mx < ok_x + btn_w && my >= btn_y && my < btn_y + btn_h {
            // OK pressed — apply will be called externally
            return true; // signal OK
        }

        // Cancel button
        let cancel_x = dx + DIALOG_W - btn_w - 16.0;
        if mx >= cancel_x && mx < cancel_x + btn_w && my >= btn_y && my < btn_y + btn_h {
            self.close();
            return true;
        }

        true // consume click (inside dialog)
    }

    /// Returns true if OK was clicked (caller should apply + close).
    pub fn is_ok_clicked(&self, mx: f32, my: f32, win_w: f32, win_h: f32) -> bool {
        if !self.visible { return false; }
        let dx = (win_w - DIALOG_W) / 2.0;
        let dy = (win_h - DIALOG_H) / 2.0;
        let footer_y = dy + DIALOG_H - FOOTER_H;
        let btn_w = 80.0;
        let btn_h = 28.0;
        let ok_x = dx + DIALOG_W - btn_w * 2.0 - 24.0;
        let btn_y = footer_y + (FOOTER_H - btn_h) / 2.0;
        mx >= ok_x && mx < ok_x + btn_w && my >= btn_y && my < btn_y + btn_h
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
