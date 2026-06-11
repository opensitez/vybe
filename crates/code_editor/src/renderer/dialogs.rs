use tiny_skia::{Paint, Pixmap, Transform};
use vybe_widgets::{FontSystem, SwashCache, TextColor as Color};

pub(crate) struct ProjectPropsDialog {
    pub visible: bool,
    pub selected_startup: usize,
}

impl ProjectPropsDialog {
    pub fn new() -> Self {
        Self {
            visible: false,
            selected_startup: 0,
        }
    }

    pub fn open(&mut self, project: &vybe_compiler::projects::project::Project) {
        self.visible = true;
        self.selected_startup = match &project.startup_object {
            vybe_compiler::projects::project::StartupObject::SubMain => 0,
            vybe_compiler::projects::project::StartupObject::None => 1,
            vybe_compiler::projects::project::StartupObject::Form(name) => project
                .forms
                .iter()
                .position(|f| &f.form.name == name)
                .map(|i| i + 2)
                .unwrap_or(0),
        };
    }

    pub fn close(&mut self) {
        self.visible = false;
    }

    pub fn apply(&self, project: &mut vybe_compiler::projects::project::Project) {
        match self.selected_startup {
            0 => {
                project.startup_object = vybe_compiler::projects::project::StartupObject::SubMain;
                project.startup_form = None;
            }
            1 => {
                project.startup_object = vybe_compiler::projects::project::StartupObject::None;
                project.startup_form = None;
            }
            n => {
                let idx = n - 2;
                if let Some(fm) = project.forms.get(idx) {
                    let name = fm.form.name.clone();
                    project.startup_object =
                        vybe_compiler::projects::project::StartupObject::Form(name.clone());
                    project.startup_form = Some(name);
                }
            }
        }
    }

    pub fn startup_options(
        &self,
        project: &vybe_compiler::projects::project::Project,
    ) -> Vec<String> {
        let mut opts = vec!["Sub Main".into(), "(None)".into()];
        for fm in &project.forms {
            opts.push(fm.form.name.clone());
        }
        opts
    }

    pub fn render(
        &self,
        pix: &mut Pixmap,
        fs: &mut FontSystem,
        sc: &mut SwashCache,
        win_w: f32,
        win_h: f32,
        scale: f32,
        project: &vybe_compiler::projects::project::Project,
    ) {
        if !self.visible {
            return;
        }
        let s = scale;
        let mut paint = Paint::default();
        let dw = 400.0f32;
        let dh = 280.0f32;
        let dx = (win_w - dw) / 2.0;
        let dy = (win_h - dh) / 2.0;

        // Overlay
        paint.set_color_rgba8(0, 0, 0, 120);
        if let Some(r) = tiny_skia::Rect::from_xywh(0.0, 0.0, win_w * s, win_h * s) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }

        // Shadow + bg
        paint.set_color_rgba8(0, 0, 0, 40);
        if let Some(r) = tiny_skia::Rect::from_xywh((dx + 4.0) * s, (dy + 4.0) * s, dw * s, dh * s)
        {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(dx * s, dy * s, dw * s, dh * s) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }

        // Title bar
        paint.set_color_rgba8(0, 120, 212, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(dx * s, dy * s, dw * s, 30.0 * s) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }
        let white = Color::rgba(255, 255, 255, 255);
        let text_col = Color::rgba(30, 30, 30, 255);
        crate::ide_text::draw_text(
            pix,
            fs,
            sc,
            &format!("{} - Project Properties", project.name),
            dx + 10.0,
            dy + 7.0,
            13.0,
            white,
            s,
        );
        crate::ide_text::draw_text(pix, fs, sc, "X", dx + dw - 22.0, dy + 7.0, 13.0, white, s);

        let label_col = Color::rgba(60, 60, 60, 255);
        let mut y = dy + 30.0 + 16.0;

        // Project Name
        crate::ide_text::draw_text(
            pix,
            fs,
            sc,
            "Project Name:",
            dx + 16.0,
            y,
            12.0,
            label_col,
            s,
        );
        y += 20.0;
        paint.set_color_rgba8(245, 245, 245, 255);
        if let Some(r) =
            tiny_skia::Rect::from_xywh((dx + 16.0) * s, y * s, (dw - 32.0) * s, 28.0 * s)
        {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }
        crate::ide_text::draw_text(
            pix,
            fs,
            sc,
            &project.name,
            dx + 22.0,
            y + 6.0,
            12.0,
            text_col,
            s,
        );
        y += 28.0 + 16.0;

        // Startup Object
        crate::ide_text::draw_text(
            pix,
            fs,
            sc,
            "Startup Object:",
            dx + 16.0,
            y,
            12.0,
            label_col,
            s,
        );
        y += 20.0;

        let options = self.startup_options(project);
        let opt_h = 22.0f32;
        for (i, label) in options.iter().enumerate() {
            let oy = y + i as f32 * opt_h;
            if i == self.selected_startup {
                paint.set_color_rgba8(0, 120, 212, 255);
                if let Some(r) = tiny_skia::Rect::from_xywh(
                    (dx + 17.0) * s,
                    (oy + 1.0) * s,
                    (dw - 34.0) * s,
                    (opt_h - 2.0) * s,
                ) {
                    pix.fill_rect(r, &paint, Transform::identity(), None);
                }
                crate::ide_text::draw_text(pix, fs, sc, label, dx + 24.0, oy + 4.0, 11.0, white, s);
            } else {
                crate::ide_text::draw_text(
                    pix,
                    fs,
                    sc,
                    label,
                    dx + 24.0,
                    oy + 4.0,
                    11.0,
                    text_col,
                    s,
                );
            }
        }

        // Footer
        let footer_y = dy + dh - 44.0;
        paint.set_color_rgba8(240, 240, 240, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(dx * s, footer_y * s, dw * s, 44.0 * s) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }

        // OK button
        let btn_w = 80.0;
        let btn_h = 28.0;
        let ok_x = dx + dw - btn_w * 2.0 - 24.0;
        let btn_y = footer_y + 8.0;
        paint.set_color_rgba8(0, 120, 212, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(ok_x * s, btn_y * s, btn_w * s, btn_h * s) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }
        crate::ide_text::draw_text(pix, fs, sc, "OK", ok_x + 30.0, btn_y + 6.0, 12.0, white, s);

        // Cancel button
        let cancel_x = dx + dw - btn_w - 16.0;
        paint.set_color_rgba8(240, 240, 240, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(cancel_x * s, btn_y * s, btn_w * s, btn_h * s) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }
        crate::ide_text::draw_text(
            pix,
            fs,
            sc,
            "Cancel",
            cancel_x + 20.0,
            btn_y + 6.0,
            12.0,
            text_col,
            s,
        );
    }

    pub fn handle_click(
        &mut self,
        mx: f32,
        my: f32,
        win_w: f32,
        win_h: f32,
        project: &vybe_compiler::projects::project::Project,
    ) -> bool {
        if !self.visible {
            return false;
        }
        let dw = 400.0f32;
        let dh = 280.0f32;
        let dx = (win_w - dw) / 2.0;
        let dy = (win_h - dh) / 2.0;

        // Outside dialog
        if mx < dx || mx > dx + dw || my < dy || my > dy + dh {
            self.close();
            return true;
        }

        // Close button
        if mx >= dx + dw - 28.0 && my < dy + 30.0 {
            self.close();
            return true;
        }

        // Startup option list
        let list_y = dy + 30.0 + 16.0 + 20.0 + 28.0 + 16.0 + 20.0;
        let options = self.startup_options(project);
        let opt_h = 22.0;
        if mx >= dx + 16.0 && mx < dx + dw - 16.0 {
            let rel_y = my - list_y;
            if rel_y >= 0.0 {
                let idx = (rel_y / opt_h) as usize;
                if idx < options.len() {
                    self.selected_startup = idx;
                    return true;
                }
            }
        }

        // Cancel button
        let footer_y = dy + dh - 44.0;
        let btn_w = 80.0;
        let btn_h = 28.0;
        let cancel_x = dx + dw - btn_w - 16.0;
        let btn_y = footer_y + 8.0;
        if mx >= cancel_x && mx < cancel_x + btn_w && my >= btn_y && my < btn_y + btn_h {
            self.close();
            return true;
        }

        true // consume click inside dialog
    }

    pub fn is_ok_clicked(&self, mx: f32, my: f32, win_w: f32, win_h: f32) -> bool {
        if !self.visible {
            return false;
        }
        let dw = 400.0f32;
        let dh = 280.0f32;
        let dx = (win_w - dw) / 2.0;
        let dy = (win_h - dh) / 2.0;
        let footer_y = dy + dh - 44.0;
        let btn_w = 80.0;
        let btn_h = 28.0;
        let ok_x = dx + dw - btn_w * 2.0 - 24.0;
        let btn_y = footer_y + 8.0;
        mx >= ok_x && mx < ok_x + btn_w && my >= btn_y && my < btn_y + btn_h
    }
}
