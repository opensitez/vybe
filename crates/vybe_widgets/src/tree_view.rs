use super::layout::{
    KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent, MouseEventKind,
    PanelWidget, RenderContext, WidgetEvent, WidgetId,
};
use cosmic_text::{Attrs, Buffer, Color as CosmicColor, Family, FontSystem, Metrics, SwashCache};
use std::fs;
use tiny_skia::{Paint, PathBuilder, Pixmap, Transform};

#[derive(Clone)]
pub struct FileEntry {
    pub name: String,
    pub path: String,
    pub is_dir: bool,
    pub is_expanded: bool,
    pub children: Vec<FileEntry>,
    pub buffer: Option<Buffer>,
}

pub enum TreeEvent {
    Open(String),
    None,
}

pub struct TreeView {
    pub entries: Vec<FileEntry>,
    pub item_height: f32,
    pub indent: f32,
    pub scale: f32,
    pub selected_path: Option<String>,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

fn scan_dir(path: &str) -> Vec<FileEntry> {
    let mut entries = Vec::new();
    if let Ok(rd) = fs::read_dir(path) {
        for e in rd.flatten() {
            let name = e.file_name().to_string_lossy().to_string();
            if name.starts_with('.') || name == "target" || name == "node_modules" {
                continue;
            }
            let p = e.path().to_string_lossy().to_string();
            let is_dir = e.file_type().map(|t| t.is_dir()).unwrap_or(false);
            entries.push(FileEntry {
                name,
                path: p,
                is_dir,
                is_expanded: false,
                children: Vec::new(),
                buffer: None,
            });
        }
    }
    entries.sort_by(|a, b| b.is_dir.cmp(&a.is_dir).then(a.name.cmp(&b.name)));
    entries
}

impl TreeView {
    pub fn new(root_path: &str, scale: f32) -> Self {
        let mut tree = Self {
            entries: Vec::new(),
            item_height: 26.0 * scale,
            indent: 20.0 * scale,
            scale,
            selected_path: None,
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        };
        tree.entries = scan_dir(root_path);
        tree
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Update scale factor (e.g. when moving to a different DPI monitor).
    /// Invalidates cached text buffers.
    pub fn set_scale(&mut self, scale: f32) {
        if (self.scale - scale).abs() < 0.01 {
            return;
        }
        self.scale = scale;
        self.item_height = 26.0 * scale;
        self.indent = 20.0 * scale;
        Self::invalidate_buffers(&mut self.entries);
    }

    fn invalidate_buffers(entries: &mut [FileEntry]) {
        for entry in entries {
            entry.buffer = None;
            Self::invalidate_buffers(&mut entry.children);
        }
    }

    pub fn reveal_path(&mut self, target_path: &str) {
        self.selected_path = Some(target_path.to_string());
        for entry in &mut self.entries {
            if Self::reveal_recursive(entry, target_path) {
                break;
            }
        }
    }

    fn reveal_recursive(entry: &mut FileEntry, target_path: &str) -> bool {
        if entry.path == target_path {
            return true;
        }
        if entry.is_dir && target_path.starts_with(&entry.path) {
            entry.is_expanded = true;
            if entry.children.is_empty() {
                entry.children = scan_dir(&entry.path);
            }
            for child in &mut entry.children {
                if Self::reveal_recursive(child, target_path) {
                    return true;
                }
            }
        }
        false
    }

    pub fn render_tree(
        &mut self,
        pix: &mut Pixmap,
        fs: &mut FontSystem,
        sc: &mut SwashCache,
        x: f32,
        y: f32,
        width: f32,
        text_color: CosmicColor,
        selection_color: (u8, u8, u8, u8),
    ) {
        let mut current_y = y;
        let scale = self.scale;
        let item_height = self.item_height;
        let indent = self.indent;

        for entry in &mut self.entries {
            Self::render_entry(
                pix,
                fs,
                sc,
                entry,
                x,
                &mut current_y,
                0,
                scale,
                item_height,
                indent,
                &self.selected_path,
                width,
                text_color,
                selection_color,
            );
        }
    }

    fn render_entry(
        pix: &mut Pixmap,
        fs: &mut FontSystem,
        sc: &mut SwashCache,
        entry: &mut FileEntry,
        x: f32,
        y: &mut f32,
        depth: u32,
        scale: f32,
        item_height: f32,
        indent: f32,
        selected_path: &Option<String>,
        width: f32,
        text_color: CosmicColor,
        selection_color: (u8, u8, u8, u8),
    ) {
        let x_off = x + (depth as f32 * indent);
        let active_y = *y;

        // 0. Draw Background for Selected
        if let Some(sel) = selected_path {
            if sel == &entry.path {
                let mut p = Paint::default();
                p.set_color_rgba8(
                    selection_color.0,
                    selection_color.1,
                    selection_color.2,
                    selection_color.3,
                );
                pix.fill_rect(
                    tiny_skia::Rect::from_xywh(x, active_y, width, item_height).unwrap(),
                    &p,
                    Transform::identity(),
                    None,
                );
            } else if entry.is_dir && sel.starts_with(&entry.path) {
                // Optional: subtle highlight for parent folders?
            }
        }
        if entry.is_dir {
            let mut pb = PathBuilder::new();
            let cx = x_off + 8.0 * scale;
            let cy = active_y + item_height / 2.0;
            let sz = 4.0 * scale;
            if entry.is_expanded {
                pb.move_to(cx - sz, cy - sz / 2.0);
                pb.line_to(cx + sz, cy - sz / 2.0);
                pb.line_to(cx, cy + sz);
            } else {
                pb.move_to(cx - sz / 2.0, cy - sz);
                pb.line_to(cx + sz, cy);
                pb.line_to(cx - sz / 2.0, cy + sz);
            }
            pb.close();
            if let Some(path) = pb.finish() {
                let mut p = Paint::default();
                p.set_color_rgba8(
                    text_color.r(),
                    text_color.g(),
                    text_color.b(),
                    text_color.a(),
                );
                pix.fill_path(
                    &path,
                    &p,
                    tiny_skia::FillRule::Winding,
                    Transform::identity(),
                    None,
                );
            }
        }

        // 1.5 Draw File/Folder Icon Dot
        let (r, g, b) = if entry.is_dir {
            (255, 214, 102) // folder yellow
        } else {
            let ext = entry.path.split('.').last().unwrap_or("");
            match ext {
                "rs" => (222, 165, 132),         // rust orange
                "js" | "ts" => (241, 224, 90),   // js yellow
                "css" | "scss" => (86, 61, 124), // css purple
                "json" => (41, 191, 191),        // json cyan
                "md" => (8, 63, 161),            // md blue
                _ => (150, 150, 150),            // generic grey
            }
        };
        let mut ic_p = Paint::default();
        ic_p.set_color_rgba8(r, g, b, 255);
        let ic_sz = 6.0 * scale;
        pix.fill_rect(
            tiny_skia::Rect::from_xywh(
                x_off + 18.0 * scale,
                active_y + (item_height - ic_sz) / 2.0,
                ic_sz,
                ic_sz,
            )
            .unwrap(),
            &ic_p,
            Transform::identity(),
            None,
        );

        // 2. Draw Name
        let text_x = x_off + 30.0 * scale;
        let text_y = active_y + 2.0 * scale;
        let col = text_color;
        if entry.buffer.is_none() {
            let mut lab = Buffer::new(fs, Metrics::new(14.0, 20.0).scale(scale));
            lab.set_text(
                fs,
                &entry.name,
                &Attrs::new().family(Family::Monospace).color(col),
                cosmic_text::Shaping::Advanced,
                None,
            );
            lab.shape_until_scroll(fs, false);
            entry.buffer = Some(lab);
        }
        if let Some(lab) = &entry.buffer {
            for r in lab.layout_runs() {
                for g in r.glyphs {
                    let pg = g.physical((text_x, text_y + r.line_y), 1.0);
                    if let Some(im) = sc.get_image(fs, pg.cache_key) {
                        let mut p =
                            Pixmap::new(im.placement.width.max(1), im.placement.height.max(1))
                                .unwrap();
                        let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a());
                        for (pix_slot, &al) in p.pixels_mut().iter_mut().zip(im.data.iter()) {
                            let af = (al as f32 / 255.0) * (ca as f32 / 255.0);
                            *pix_slot = tiny_skia::ColorU8::from_rgba(
                                (cr as f32 * af) as u8,
                                (cg as f32 * af) as u8,
                                (cb as f32 * af) as u8,
                                (255.0 * af) as u8,
                            )
                            .premultiply();
                        }
                        pix.draw_pixmap(
                            pg.x + im.placement.left,
                            pg.y - im.placement.top,
                            p.as_ref(),
                            &tiny_skia::PixmapPaint::default(),
                            Transform::identity(),
                            None,
                        );
                    }
                }
            }
        }

        *y += item_height;

        // 3. Recursive children
        if entry.is_expanded && entry.is_dir {
            if entry.children.is_empty() {
                entry.children = scan_dir(&entry.path);
            }
            for child in &mut entry.children {
                Self::render_entry(
                    pix,
                    fs,
                    sc,
                    child,
                    x,
                    y,
                    depth + 1,
                    scale,
                    item_height,
                    indent,
                    selected_path,
                    width,
                    text_color,
                    selection_color,
                );
            }
        }
    }

    pub fn handle_mouse_at(&mut self, mx: f32, my: f32, x: f32, y: f32) -> TreeEvent {
        let mut current_y = y;
        let mut result = TreeEvent::None;
        let item_height = self.item_height;
        let indent = self.indent;

        for entry in &mut self.entries {
            if let Some(ev) =
                Self::check_mouse_entry(entry, mx, my, x, &mut current_y, 0, item_height, indent)
            {
                result = ev;
                break;
            }
        }
        result
    }

    fn check_mouse_entry(
        entry: &mut FileEntry,
        mx: f32,
        my: f32,
        x: f32,
        y: &mut f32,
        depth: u32,
        item_height: f32,
        indent: f32,
    ) -> Option<TreeEvent> {
        let _x_off = x + (depth as f32 * indent);
        let active_y = *y;

        if my >= active_y && my < active_y + item_height {
            if entry.is_dir {
                entry.is_expanded = !entry.is_expanded;
                return Some(TreeEvent::None);
            } else {
                return Some(TreeEvent::Open(entry.path.clone()));
            }
        }
        *y += item_height;
        if entry.is_expanded && entry.is_dir {
            for child in &mut entry.children {
                if let Some(ev) =
                    Self::check_mouse_entry(child, mx, my, x, y, depth + 1, item_height, indent)
                {
                    return Some(ev);
                }
            }
        }
        None
    }

    pub fn draw_text_static_internal(
        pix: &mut Pixmap,
        fs: &mut FontSystem,
        sc: &mut SwashCache,
        text: &str,
        x: f32,
        y: f32,
        col: CosmicColor,
        scale: f32,
    ) {
        let mut lab = Buffer::new(fs, Metrics::new(14.0, 20.0).scale(scale));
        lab.set_text(
            fs,
            text,
            &Attrs::new().family(Family::Monospace).color(col),
            cosmic_text::Shaping::Advanced,
            None,
        );
        lab.shape_until_scroll(fs, false);
        for r in lab.layout_runs() {
            for g in r.glyphs {
                let pg = g.physical((x, y + r.line_y), 1.0);
                if let Some(im) = sc.get_image(fs, pg.cache_key) {
                    let mut p =
                        Pixmap::new(im.placement.width.max(1), im.placement.height.max(1)).unwrap();
                    let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a());
                    for (pix_slot, &al) in p.pixels_mut().iter_mut().zip(im.data.iter()) {
                        let af = (al as f32 / 255.0) * (ca as f32 / 255.0);
                        *pix_slot = tiny_skia::ColorU8::from_rgba(
                            (cr as f32 * af) as u8,
                            (cg as f32 * af) as u8,
                            (cb as f32 * af) as u8,
                            (255.0 * af) as u8,
                        )
                        .premultiply();
                    }
                    pix.draw_pixmap(
                        pg.x + im.placement.left,
                        pg.y - im.placement.top,
                        p.as_ref(),
                        &tiny_skia::PixmapPaint::default(),
                        Transform::identity(),
                        None,
                    );
                }
            }
        }
    }
}

impl PanelWidget for TreeView {
    fn name(&self) -> &str {
        &self.name
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
    }
    fn rect(&self) -> LayoutRect {
        self.rect
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }
        self.set_scale(ctx.scale);
        let text_color = CosmicColor::rgba(200, 200, 200, 255);
        let selection_color: (u8, u8, u8, u8) = (0, 120, 215, 80);
        self.render_tree(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            r.x,
            r.y,
            r.w,
            text_color,
            selection_color,
        );
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        if !r.contains(event.x, event.y) {
            return false;
        }
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            match self.handle_mouse_at(event.x, event.y, r.x, r.y) {
                TreeEvent::Open(path) => {
                    self.selected_path = Some(path.clone());
                    self.pending_events.push(WidgetEvent::TreeItemOpened(path));
                    return true;
                }
                TreeEvent::None => {
                    return true;
                }
            }
        }
        false
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }
    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }
}
