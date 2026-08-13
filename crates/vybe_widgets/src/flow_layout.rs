//! FlowLayoutPanel — arranges children in a flow (left-to-right, wrapping to next row).
//!
//! Like WinForms FlowLayoutPanel: children are placed sequentially; when a child
//! would overflow the current row it wraps to the next row.

use super::WidgetColors;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext, WidgetCommand,
    WidgetEvent, WidgetId, command_color, command_number,
};
use tiny_skia::*;

/// Flow direction for child arrangement.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FlowDirection {
    LeftToRight,
    TopDown,
}

pub struct FlowLayoutPanel {
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    pub flow_direction: FlowDirection,
    /// Spacing between children in pixels.
    pub spacing: f32,
    /// Padding inside the panel edges.
    pub padding: f32,
    /// Whether children wrap to the next row/column when they exceed the panel size.
    pub wrap_contents: bool,
    /// This panel's flex weight when it is itself a child of another flex
    /// container (0 = fixed/natural size). Default 1.
    pub flex: f32,
    /// Fixed main-axis size (height in TopDown, width in LeftToRight) used when
    /// `flex == 0`. Default ~a toolbar height.
    pub fixed_main: f32,
    rect: LayoutRect,
    children: Vec<Box<dyn PanelWidget>>,
    /// `justify-content` — how leftover MAIN-axis space is distributed.
    ///
    /// Only observable when something is left over, which means when no child
    /// grows: a flex child already eats the remainder, and there is then
    /// nothing to distribute. Default `flex-start` is what this panel has
    /// always done, so declaring nothing changes nothing.
    pub justify_content: String,
    /// `align-items` — the CROSS-axis size of each child. `stretch` (the
    /// default, and the panel's long-standing behaviour) fills the container;
    /// anything else leaves the child its natural size and places it.
    pub align_items: String,
    /// Per-child `flex-grow`, by child name.
    ///
    /// Here rather than on the child because `layout_flex()` is a TRAIT method
    /// only containers implement — a button or a label falls through to the
    /// default of `1.0`, so `SetFlex` addressed to a leaf did nothing and every
    /// leaf grew. Flutter's `Expanded` is precisely the distinction that was
    /// being lost. Same reasoning as `out_of_flow`: a fact about how the
    /// container treats a child lives with the container.
    child_flex: std::collections::HashMap<String, f32>,
    /// Each child's `(width, height)` as it arrived, before this panel resized
    /// it. `align-items` needs a size to align *within*, and once a child has
    /// been stretched its rect no longer remembers one.
    natural: std::collections::HashMap<String, (f32, f32)>,
    /// Per-child `align-self` — one child overruling `align_items`.
    child_align: std::collections::HashMap<String, String>,
    /// Per-child `order`. Children are arranged by this before document order,
    /// which is what CSS does; absent means `0`.
    child_order: std::collections::HashMap<String, i32>,
    /// `overflow` — whether content outside this panel is clipped.
    pub overflow: String,
    /// Draw a border with a caption gap — a `<fieldset>`, not a `<div>`.
    pub bordered: bool,
    /// The caption drawn across the top edge — HTML's `<legend>`, VCL's
    /// `TGroupBox.Caption`.
    pub caption: String,
    /// Children this panel must NOT arrange — CSS `position: absolute`.
    ///
    /// The inverse of `dock`, and here for the same reason it is: a docked
    /// child does not own its position, so the container computes it; an
    /// out-of-flow child DOES own its position, so the container has to be told
    /// to leave it alone. Neither fact belongs to the child's widget type,
    /// which is why both live with whoever does the arranging.
    ///
    /// Without this, `relayout()` recomputed every child's rect from flow
    /// order and discarded whatever `left`/`top` had put there. A control was
    /// not *missing* a rect — it had the wrong one, which is why it rendered
    /// in flow order rather than at nonsense coordinates.
    out_of_flow: std::collections::HashSet<String>,
    /// `position: relative` children, and how far each is offset.
    ///
    /// The OTHER half of positioning, and deliberately not the same set as
    /// `out_of_flow`: a relative box KEEPS its flow slot — its siblings do not
    /// close up behind it — and is merely drawn offset from where the flow put
    /// it. Only `absolute`/`fixed` leave the flow. Treating the two alike is
    /// what left `relative` with no way to mean what CSS says it means.
    relative_offset: std::collections::HashMap<String, (f32, f32)>,
}

impl FlowLayoutPanel {
    pub fn new() -> Self {
        Self {
            width: 300.0,
            height: 200.0,
            colors: WidgetColors {
                background: (250, 250, 250, 255),
                border: (180, 180, 180, 255),
                ..WidgetColors::default()
            },
            id: WidgetId::next(),
            name: String::new(),
            flow_direction: FlowDirection::TopDown,
            spacing: 4.0,
            padding: 4.0,
            wrap_contents: true,
            flex: 1.0,
            fixed_main: 44.0,
            rect: LayoutRect::zero(),
            children: Vec::new(),
            justify_content: "flex-start".to_string(),
            align_items: "stretch".to_string(),
            child_flex: std::collections::HashMap::new(),
            natural: std::collections::HashMap::new(),
            child_align: std::collections::HashMap::new(),
            child_order: std::collections::HashMap::new(),
            overflow: "visible".to_string(),
            bordered: false,
            caption: String::new(),
            out_of_flow: std::collections::HashSet::new(),
            relative_offset: std::collections::HashMap::new(),
        }
    }

    /// The order children are arranged in: `order` first, document order
    /// within it. Returns indices into `self.children`.
    ///
    /// A stable sort, so children sharing an `order` keep the order they were
    /// added in — which is what CSS specifies and what makes the default (no
    /// `order` anywhere) exactly document order.
    fn arrangement(&self) -> Vec<usize> {
        let mut indices: Vec<usize> = (0..self.children.len()).collect();
        if self.child_order.is_empty() {
            return indices;
        }
        indices.sort_by_key(|i| {
            self.child_order
                .get(self.children[*i].name())
                .copied()
                .unwrap_or(0)
        });
        indices
    }

    /// Leading offset and inter-child gap for `justify-content`, given how much
    /// main-axis space is left over after the children are sized.
    ///
    /// Returns `(leading, extra_gap)`. `flex-start` is `(0, 0)`, which is what
    /// this panel did before the property existed.
    fn justify(&self, leftover: f32, count: usize) -> (f32, f32) {
        let leftover = leftover.max(0.0);
        let n = count as f32;
        match self.justify_content.as_str() {
            "flex-end" | "end" => (leftover, 0.0),
            "center" => (leftover / 2.0, 0.0),
            "space-between" if count > 1 => (0.0, leftover / (n - 1.0)),
            "space-around" if count > 0 => {
                let gap = leftover / n;
                (gap / 2.0, gap)
            }
            "space-evenly" if count > 0 => {
                let gap = leftover / (n + 1.0);
                (gap, gap)
            }
            _ => (0.0, 0.0),
        }
    }

    /// How far down the content starts.
    ///
    /// A `<legend>` occupies vertical space — content in a `<fieldset>` begins
    /// BELOW the caption, not behind it. Without this the caption drew across
    /// the first child.
    fn top_inset(&self) -> f32 {
        if self.bordered && !self.caption.is_empty() {
            self.padding + 14.0
        } else {
            self.padding
        }
    }

    /// Does this panel arrange `child`, or does the child place itself?
    fn arranges(&self, child: &dyn PanelWidget) -> bool {
        !self.out_of_flow.contains(child.name())
    }

    /// A child's grow weight — a declared `flex` if there is one, otherwise
    /// whatever the child says about itself.
    fn flex_of(&self, child: &dyn PanelWidget) -> f32 {
        self.child_flex
            .get(child.name())
            .copied()
            .unwrap_or_else(|| child.layout_flex())
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }
    pub fn with_direction(mut self, dir: FlowDirection) -> Self {
        self.flow_direction = dir;
        self
    }
    pub fn with_spacing(mut self, spacing: f32) -> Self {
        self.spacing = spacing;
        self
    }
    pub fn with_padding(mut self, padding: f32) -> Self {
        self.padding = padding;
        self
    }
    pub fn with_wrap(mut self, wrap: bool) -> Self {
        self.wrap_contents = wrap;
        self
    }

    /// Add a child widget. Triggers relayout.
    pub fn add(&mut self, widget: Box<dyn PanelWidget>) {
        // The child's own size, recorded BEFORE the first layout pass — after
        // it, the rect is whatever the panel stretched it to, and `align-items`
        // would have nothing smaller to fall back to. There is no intrinsic-size
        // query on `PanelWidget`, so insertion is the one moment the natural
        // size is still visible.
        let rect = widget.rect();
        self.natural
            .insert(widget.name().to_string(), (rect.w, rect.h));
        self.children.push(widget);
        self.relayout();
    }

    pub fn child(&self, index: usize) -> &dyn PanelWidget {
        &*self.children[index]
    }
    pub fn child_mut(&mut self, index: usize) -> &mut dyn PanelWidget {
        &mut *self.children[index]
    }
    pub fn child_count(&self) -> usize {
        self.children.len()
    }

    /// Remove a child by index. Triggers relayout.
    pub fn remove(&mut self, index: usize) -> Box<dyn PanelWidget> {
        let w = self.children.remove(index);
        self.relayout();
        w
    }

    /// Arrange children according to flow direction.
    fn relayout(&mut self) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }

        match self.flow_direction {
            FlowDirection::LeftToRight => self.layout_left_to_right(),
            FlowDirection::TopDown => self.layout_top_down(),
        }
    }

    // Flutter Row: children side by side. Fixed-flex children keep their
    // `fixed_main` width; flex children share the leftover by weight. Full
    // height.
    fn layout_left_to_right(&mut self) {
        let r = self.rect;
        let n = self.in_flow_count();
        if n == 0 {
            return;
        }
        let inner_w = (r.w - 2.0 * self.padding).max(0.0);
        let inner_h = (r.h - self.top_inset() - self.padding).max(0.0);
        let gaps = self.spacing * (n as f32 - 1.0);
        let (total_flex, fixed) = self.flex_totals();
        let leftover = (inner_w - gaps - fixed).max(0.0);
        // Nothing grows → the leftover is real free space, and `justify-content`
        // is what decides where it goes. With a growing child there is nothing
        // left to distribute, which is why the two never fight.
        let (lead, extra) = if total_flex > 0.0 {
            (0.0, 0.0)
        } else {
            self.justify(leftover, n)
        };
        let mut cx = r.x + self.padding + lead;
        let cy = r.y + self.top_inset();
        let out_of_flow = self.out_of_flow.clone();
        let relative = self.relative_offset.clone();
        let child_flex = self.child_flex.clone();
        let natural_sizes = self.natural.clone();
        let child_align = self.child_align.clone();
        let (align, spacing) = (self.align_items.clone(), self.spacing);
        for i in self.arrangement() {
            let name = self.children[i].name().to_string();
            // Out of flow: the child keeps the rect its own `left`/`top` gave
            // it, and takes no space from its siblings.
            if out_of_flow.contains(&name) {
                continue;
            }
            let child = &mut self.children[i];
            let f = child_flex
                .get(&name)
                .copied()
                .unwrap_or_else(|| child.layout_flex());
            let cw = if f <= 0.0 {
                Self::child_fixed(child.as_ref())
            } else if total_flex > 0.0 {
                leftover * f / total_flex
            } else {
                0.0
            };
            let natural = natural_sizes
                .get(&name)
                .map(|(_, h)| *h)
                .unwrap_or_else(|| child.rect().h)
                .max(1.0);
            let mode = child_align.get(&name).unwrap_or(&align);
            let (offset, ch) = Self::align_with(mode, inner_h, natural);
            // The flow slot, then the relative offset on top of it. `cx`
            // advances by the SLOT, so offsetting a child never moves its
            // siblings — the half that distinguishes relative from absolute.
            let (dx, dy) = relative.get(&name).copied().unwrap_or((0.0, 0.0));
            child.set_rect(LayoutRect::new(cx + dx, cy + offset + dy, cw, ch));
            cx += cw + spacing + extra;
        }
    }

    /// `align` as a free function, so a child loop already holding `&mut self`
    /// can still ask.
    fn align_with(align_items: &str, inner_cross: f32, natural: f32) -> (f32, f32) {
        match align_items {
            "flex-start" | "start" => (0.0, natural.min(inner_cross)),
            "flex-end" | "end" => {
                let size = natural.min(inner_cross);
                (inner_cross - size, size)
            }
            "center" => {
                let size = natural.min(inner_cross);
                ((inner_cross - size) / 2.0, size)
            }
            _ => (0.0, inner_cross),
        }
    }

    // Flutter Column: children stacked top to bottom. Fixed-flex children keep
    // their `fixed_main` height; flex children share the leftover by weight.
    // Full width.
    fn layout_top_down(&mut self) {
        let r = self.rect;
        let n = self.in_flow_count();
        if n == 0 {
            return;
        }
        let inner_w = (r.w - 2.0 * self.padding).max(0.0);
        let inner_h = (r.h - self.top_inset() - self.padding).max(0.0);
        let gaps = self.spacing * (n as f32 - 1.0);
        let (total_flex, fixed) = self.flex_totals();
        let leftover = (inner_h - gaps - fixed).max(0.0);
        let (lead, extra) = if total_flex > 0.0 {
            (0.0, 0.0)
        } else {
            self.justify(leftover, n)
        };
        let cx = r.x + self.padding;
        let mut cy = r.y + self.top_inset() + lead;
        let out_of_flow = self.out_of_flow.clone();
        let relative = self.relative_offset.clone();
        let child_flex = self.child_flex.clone();
        let natural_sizes = self.natural.clone();
        let child_align = self.child_align.clone();
        let (align, spacing) = (self.align_items.clone(), self.spacing);
        for i in self.arrangement() {
            let name = self.children[i].name().to_string();
            if out_of_flow.contains(&name) {
                continue;
            }
            let child = &mut self.children[i];
            let f = child_flex
                .get(&name)
                .copied()
                .unwrap_or_else(|| child.layout_flex());
            let ch = if f <= 0.0 {
                Self::child_fixed(child.as_ref())
            } else if total_flex > 0.0 {
                leftover * f / total_flex
            } else {
                0.0
            };
            let natural = natural_sizes
                .get(&name)
                .map(|(w, _)| *w)
                .unwrap_or_else(|| child.rect().w)
                .max(1.0);
            let mode = child_align.get(&name).unwrap_or(&align);
            let (offset, cw) = Self::align_with(mode, inner_w, natural);
            let (dx, dy) = relative.get(&name).copied().unwrap_or((0.0, 0.0));
            child.set_rect(LayoutRect::new(cx + offset + dx, cy + dy, cw, ch));
            cy += ch + spacing + extra;
        }
    }

    /// How many children this panel actually arranges.
    fn in_flow_count(&self) -> usize {
        self.children
            .iter()
            .filter(|c| self.arranges(c.as_ref()))
            .count()
    }

    /// (sum of flex weights, sum of fixed children's main-axis sizes).
    ///
    /// Out-of-flow children contribute NOTHING — they take no space from their
    /// siblings, which is the half of `position: absolute` that is easy to
    /// forget and shows up as everything else being squeezed.
    fn flex_totals(&self) -> (f32, f32) {
        let mut total_flex = 0.0;
        let mut fixed = 0.0;
        for child in &self.children {
            if !self.arranges(child.as_ref()) {
                continue;
            }
            let f = self.flex_of(child.as_ref());
            if f <= 0.0 {
                fixed += Self::child_fixed(child.as_ref());
            } else {
                total_flex += f;
            }
        }
        (total_flex, fixed)
    }

    /// The fixed main-axis size of a flex-0 child (a toolbar-height bar).
    fn child_fixed(_child: &dyn PanelWidget) -> f32 {
        44.0
    }

    /// Paint — a Flutter layout container is transparent (no chrome); only its
    /// children paint. Kept as a no-op so Column/Row/Scaffold don't draw the
    /// WinForms-style dashed panel border.
    /// A layout panel paints nothing — a `<div>` has no border in a browser
    /// either, and inventing one would be wrong.
    ///
    /// A `<fieldset>` DOES: the UA stylesheet gives it a groove border, and its
    /// `<legend>` sits in a gap at the top. That is why this is a flag rather
    /// than always-on — the same widget backs both elements, and only one of
    /// them is supposed to draw.
    ///
    /// Without it a group box rendered **nothing at all** — no border, no
    /// caption — while positioning its children perfectly, which is a
    /// confusing failure to look at: the container is invisible and its
    /// contents are exactly where they should be.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        if !self.bordered {
            return;
        }
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let stroke = Stroke {
            width: 1.0,
            ..Stroke::default()
        };
        let (r, g, b, a) = self.colors.border;
        paint.set_color_rgba8(r, g, b, a);

        // The top edge is dropped so the caption can sit across it, and broken
        // where the caption goes — the `<legend>` gap.
        let top = y + 8.0;
        let (right, bottom) = (x + self.rect.w, y + self.rect.h);
        let caption_w = if self.caption.is_empty() {
            0.0
        } else {
            // Rough advance; the caption is drawn by the text layer.
            self.caption.chars().count() as f32 * 7.0
        };
        let gap_start = x + 8.0;
        let gap_end = gap_start + caption_w + if caption_w > 0.0 { 8.0 } else { 0.0 };

        let mut edge = |x0: f32, y0: f32, x1: f32, y1: f32| {
            let mut pb = PathBuilder::new();
            pb.move_to(x0, y0);
            pb.line_to(x1, y1);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
        };
        edge(x, top, gap_start, top);
        edge(gap_end, top, right, top);
        edge(right, top, right, bottom);
        edge(right, bottom, x, bottom);
        edge(x, bottom, x, top);
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}

impl PanelWidget for FlowLayoutPanel {
    fn name(&self) -> &str {
        &self.name
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }

    /// The document tree's children — what `find_widget_mut` / `take_widget`
    /// walk, and what makes a node reachable by name however deeply nested.
    fn children_mut(&mut self) -> Vec<&mut Box<dyn PanelWidget>> {
        self.children.iter_mut().collect()
    }

    /// `removeChild` against a direct child.
    fn detach(&mut self, name: &str) -> Option<Box<dyn PanelWidget>> {
        let i = self.children.iter().position(|c| c.name() == name)?;
        Some(self.children.remove(i))
    }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.width = rect.w;
        self.height = rect.h;
        self.relayout();
    }
    fn rect(&self) -> LayoutRect {
        self.rect
    }

    fn add_child(&mut self, child: Box<dyn PanelWidget>) -> Option<Box<dyn PanelWidget>> {
        self.add(child); // pushes + relayout()
        None
    }

    fn send_command_named(&mut self, name: &str, cmd: &WidgetCommand) -> Option<CommandValue> {
        if self.name == name {
            return Some(self.handle_command(cmd));
        }
        for child in &mut self.children {
            if let Some(result) = child.send_command_named(name, cmd) {
                return Some(result);
            }
        }
        None
    }

    fn add_child_to(
        &mut self,
        parent_name: &str,
        child: Box<dyn PanelWidget>,
    ) -> Option<Box<dyn PanelWidget>> {
        if self.name == parent_name {
            return self.add_child(child);
        }
        let mut child = Some(child);
        for existing in &mut self.children {
            if let Some(c) = child.take() {
                child = existing.add_child_to(parent_name, c);
            }
            if child.is_none() {
                break;
            }
        }
        child
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
        // The caption sits IN the gap the border left for it — HTML's
        // `<legend>`. `paint` only has a pixmap; text needs the render context,
        // so it is drawn here rather than there.
        if self.bordered && !self.caption.is_empty() {
            let (tr, tg, tb, ta) = self.colors.foreground;
            ctx.draw_text(&self.caption, r.x + 12.0, r.y + 2.0, tr, tg, tb, ta);
        }
        for child in &mut self.children {
            child.render(ctx);
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        for child in self.children.iter_mut().rev() {
            if child.rect().contains(event.x, event.y) && child.handle_mouse(event) {
                return true;
            }
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        for child in &mut self.children {
            if child.handle_key(event) {
                return true;
            }
        }
        false
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        for child in self.children.iter_mut().rev() {
            if child.rect().contains(x, y) && child.handle_scroll(delta, x, y) {
                return true;
            }
        }
        false
    }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        for child in self.children.iter().rev() {
            if child.rect().contains(x, y) {
                return child.cursor_at(x, y);
            }
        }
        winit::window::CursorIcon::Default
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        let mut events = Vec::new();
        for child in &mut self.children {
            events.extend(child.drain_events());
        }
        events
    }

    fn layout_flex(&self) -> f32 {
        self.flex
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetFlex(f) => {
                self.flex = *f;
                CommandValue::None
            }
            // CSS layout, addressed to the container because the container is
            // what arranges. `SetChildFlow` names a child and says which of the
            // three placements it has: this panel arranges it (`static`), it
            // keeps its own coordinates (`absolute`/`fixed`), or this panel
            // arranges it and THEN offsets it (`relative:dx,dy`).
            WidgetCommand::Custom(name, value) if name == "SetChildFlow" => {
                if let CommandValue::Text(spec) = value {
                    match spec.split_once('=') {
                        Some((child, "flow")) => {
                            self.out_of_flow.remove(child);
                            self.relative_offset.remove(child);
                        }
                        // `relative` is IN flow — the container still arranges
                        // it, and the offset is applied to the slot it was
                        // given. So it never enters `out_of_flow`, and its
                        // siblings keep sitting where they would if the offset
                        // were zero. That is the whole difference from
                        // `absolute`, and the reason both are needed.
                        Some((child, rest)) if rest.starts_with("relative:") => {
                            self.out_of_flow.remove(child);
                            let offsets = &rest["relative:".len()..];
                            let (dx, dy) = offsets.split_once(',').unwrap_or(("0", "0"));
                            self.relative_offset.insert(
                                child.to_string(),
                                (
                                    dx.trim().parse().unwrap_or(0.0),
                                    dy.trim().parse().unwrap_or(0.0),
                                ),
                            );
                        }
                        Some((child, _)) => {
                            // Only on the TRANSITION into out-of-flow. Being
                            // told again is not new information, and restoring
                            // on every assertion undid work done in between —
                            // `right` computes a width from `left`, then
                            // re-asserts out-of-flow, and the restore threw the
                            // width away.
                            let newly = self.out_of_flow.insert(child.to_string());
                            // Give the child its natural SIZE back.
                            //
                            // By the time a child is excluded, this panel has
                            // already laid it out at least once — insertion
                            // runs a pass — so its rect carries a stretched
                            // width or height that the flow chose. The caller
                            // then restores the axes the program DECLARED, and
                            // any axis it did not declare would keep the
                            // stretched value forever.
                            //
                            // That is why a control declaring `Left`/`Top` and
                            // a `Width` but no `Height` rendered as a tall box
                            // spanning its container, while its sibling that
                            // declared both behaved: the bug was never about
                            // position, it was the undeclared axis.
                            if let Some((w, h)) = self.natural.get(child).copied().filter(|_| newly)
                            {
                                if let Some(widget) =
                                    self.children.iter_mut().find(|c| c.name() == child)
                                {
                                    let r = widget.rect();
                                    widget.set_rect(LayoutRect::new(r.x, r.y, w, h));
                                }
                            }
                        }
                        None => {}
                    }
                    self.relayout();
                }
                CommandValue::None
            }
            // `flex-direction` — the axis children are arranged along.
            WidgetCommand::Custom(name, value) if name == "SetFlexDirection" => {
                if let CommandValue::Text(direction) = value {
                    self.flow_direction = match direction.as_str() {
                        "row" | "row-reverse" => FlowDirection::LeftToRight,
                        _ => FlowDirection::TopDown,
                    };
                    self.relayout();
                }
                CommandValue::None
            }
            // `gap` and `padding`. Both already existed as panel fields with no
            // route in from CSS, which is the whole shape of this work: the
            // layout algorithms were here, the vocabulary could not reach them.
            WidgetCommand::Custom(name, value) if name == "SetGap" => {
                if let Some(gap) = command_number(value) {
                    self.spacing = gap as f32;
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetPadding" => {
                if let Some(padding) = command_number(value) {
                    self.padding = padding as f32;
                    self.relayout();
                }
                CommandValue::None
            }
            // Per-child `flex`, addressed to the container. `SetFlex` on a leaf
            // does nothing — only panels implement `layout_flex` — so a
            // declared weight has to be recorded by whoever arranges.
            WidgetCommand::Custom(name, value) if name == "SetChildFlex" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, weight)) = spec.rsplit_once('=') {
                        if let Ok(weight) = weight.parse::<f32>() {
                            self.child_flex.insert(child.to_string(), weight);
                            self.relayout();
                        }
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetChildAlignSelf" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, mode)) = spec.rsplit_once('=') {
                        self.child_align.insert(child.to_string(), mode.to_string());
                        self.relayout();
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetChildOrder" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, order)) = spec.rsplit_once('=') {
                        if let Ok(order) = order.parse::<i32>() {
                            self.child_order.insert(child.to_string(), order);
                            self.relayout();
                        }
                    }
                }
                CommandValue::None
            }
            // A group box's caption. A layout panel has no text of its own, so
            // this was dropped entirely — `TGroupBox.Caption := 'Group'` set
            // nothing and reported nothing.
            WidgetCommand::SetText(text) => {
                self.caption = text.clone();
                CommandValue::None
            }
            WidgetCommand::GetText => CommandValue::Text(self.caption.clone()),
            WidgetCommand::Custom(name, value) if name == "SetBordered" => {
                self.bordered = !matches!(value, CommandValue::Bool(false));
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetOverflow" => {
                if let CommandValue::Text(mode) = value {
                    self.overflow = mode.clone();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetJustifyContent" => {
                if let CommandValue::Text(mode) = value {
                    self.justify_content = mode.clone();
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetAlignItems" => {
                if let CommandValue::Text(mode) = value {
                    self.align_items = mode.clone();
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetFlexWrap" => {
                if let CommandValue::Text(wrap) = value {
                    self.wrap_contents = wrap != "nowrap";
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::SetEnabled(_) | WidgetCommand::SetVisible(_) => {
                for child in &mut self.children {
                    child.handle_command(cmd);
                }
                CommandValue::None
            }
            // The panel already carries padding/spacing/size/colour; these
            // expose them so a container adapter (Flutter `Padding`/`Container`/
            // `SizedBox`, WinForms `Panel.Padding`) can drive them.
            WidgetCommand::Custom(key, val) => match key.as_str() {
                "SetPadding" => {
                    if let Some(p) = command_number(val) {
                        self.padding = p as f32;
                    }
                    CommandValue::None
                }
                "SetSpacing" => {
                    if let Some(s) = command_number(val) {
                        self.spacing = s as f32;
                    }
                    CommandValue::None
                }
                "SetWidth" => {
                    if let Some(w) = command_number(val) {
                        self.width = w as f32;
                        // A container given an explicit size is no longer
                        // free to absorb leftover space.
                        self.flex = 0.0;
                        self.fixed_main = w as f32;
                    }
                    CommandValue::None
                }
                "SetHeight" => {
                    if let Some(h) = command_number(val) {
                        self.height = h as f32;
                        self.flex = 0.0;
                        self.fixed_main = h as f32;
                    }
                    CommandValue::None
                }
                "SetBackColor" => {
                    if let Some(rgba) = command_color(val) {
                        self.colors.background = rgba;
                    }
                    CommandValue::None
                }
                _ => CommandValue::None,
            },
            _ => CommandValue::None,
        }
    }
}
