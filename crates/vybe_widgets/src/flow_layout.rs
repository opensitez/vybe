//! FlowLayoutPanel — arranges children in a flow (left-to-right, wrapping to next row).
//!
//! Like WinForms FlowLayoutPanel: children are placed sequentially; when a child
//! would overflow the current row it wraps to the next row.

use super::WidgetColors;
use crate::css::Edges;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext, WidgetCommand,
    WidgetEvent, WidgetId, command_color, command_edges, command_number,
};
use tiny_skia::*;

/// Flow direction for child arrangement.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FlowDirection {
    LeftToRight,
    TopDown,
}

/// **Which formatting context a box establishes for its children.**
///
/// The most consequential fact CSS has about a container, and the one this
/// panel did not have. Every container ran [`Formatting::Flex`] — so a `<div>`,
/// which is `display: block`, was laid out as `display: flex; flex-direction:
/// column; align-items: stretch`. Those are different CSS values that nobody
/// wrote. It is why a row of `<button>`s came out as a column of full-width
/// bars: `stretch` set each child's cross size to the container's, and the flex
/// weights divided the container's height between them.
///
/// The child's own `display` was computed correctly the whole time and simply
/// had no reader. `set_style_property`'s `display` arm said so out loud —
/// *"`block` leaves it as it was"* — which is only harmless if what it was left
/// as is block, and it was flex.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Formatting {
    /// `display: flex`. What this panel has always done, and what Flutter's
    /// `Row`/`Column` mean — they reach this widget by KIND through `vybe:gui`
    /// and never become elements, so this stays their behaviour exactly.
    Flex,
    /// **CSS normal flow** — block-level children stack, inline-level children
    /// share line boxes and wrap at the content edge.
    Normal,
    /// `display: grid`. Items are placed into the cells of a track template.
    Grid,
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
    /// Padding inside the panel edges, per side.
    ///
    /// Four values, not one, because CSS has four and a uniform scalar cannot
    /// spell `padding-left: 40px` — which is what a `<ul>` is. Children are
    /// arranged inside the **content box**, so this is what separates the rect
    /// the panel occupies from the area it hands out.
    pub padding: Edges,
    /// Whether children wrap to the next row/column when they exceed the panel size.
    ///
    /// **Defaults to `false`, which is CSS's `flex-wrap: nowrap`.** It defaulted
    /// to `true` before, but nothing read the field — wrapping was accepted and
    /// never implemented — so the old default was unobservable and no caller can
    /// have depended on it. Now that lines really break, the default has to be
    /// the one CSS specifies or every existing flex container silently changes
    /// shape.
    pub wrap_contents: bool,
    /// `row-gap` and `column-gap`, each `None` until declared.
    ///
    /// **Two axes, because CSS has two.** All three of `gap`, `row-gap` and
    /// `column-gap` used to collapse onto the single `spacing` scalar, so a box
    /// asking for `row-gap: 20px; column-gap: 4px` got whichever arrived last on
    /// both axes. `gap` sets both; the longhands set one each.
    ///
    /// `None` falls back to [`FlowLayoutPanel::spacing`], which is the toolkit's
    /// own inter-child spacing and predates any of this — so a panel nobody has
    /// spoken to keeps spacing exactly as it was.
    row_gap: Option<f32>,
    column_gap: Option<f32>,
    /// `row-reverse` / `column-reverse` — the main axis runs backwards.
    ///
    /// Both used to fold into their forward direction (`"row" | "row-reverse"
    /// => LeftToRight`), so the declaration was accepted and the order was
    /// never reversed. Flutter reaches this through
    /// `VerticalDirection.up` and a right-to-left `Row`.
    reverse: bool,
    /// `align-content` — how wrapped LINES are distributed across the cross
    /// axis. `flex-start` is the default and what a single line always gets.
    align_content: String,
    /// Per-child `grid-column` / `grid-row`, as
    /// `(column start, column end, row start, row end)`.
    child_grid_area: std::collections::HashMap<String, (crate::css::GridLine, crate::css::GridLine, crate::css::GridLine, crate::css::GridLine)>,
    /// `grid-template-columns` / `grid-template-rows`.
    ///
    /// Empty means no template, which is not the same as a grid with no
    /// columns: CSS gives an untemplated grid ONE implicit full-width column,
    /// so items stack. That default lives in `layout_grid`, not here, because
    /// "the author said nothing" and "the author said one column" are different
    /// facts and only one of them survives a round trip.
    grid_columns: Vec<crate::css::TrackSize>,
    grid_rows: Vec<crate::css::TrackSize>,
    /// Per-child `flex-basis`, in pixels — the item's size BEFORE any growing
    /// or shrinking. Both this and `child_shrink` parsed into `css.rs` and had
    /// no arm in `dom.rs`, so they were stored and never asked for.
    child_basis: std::collections::HashMap<String, f32>,
    /// Per-child `flex-shrink`.
    ///
    /// ⚠ CSS's initial value is `1`; absent here means **no shrinking**, which
    /// is what this panel has always done. Applying the real initial value
    /// would resize every existing Flutter `Row` whose children overflow, so
    /// only an explicit declaration shrinks. Recorded as a divergence.
    child_shrink: std::collections::HashMap<String, f32>,
    /// Children whose `width` was DECLARED, so normal flow must not stretch
    /// them. Absent means `auto`, which is CSS's initial value and the case
    /// where filling the content width is correct.
    child_declared_width: std::collections::HashSet<String>,
    /// Children whose `height` was DECLARED. Grid stretches an item to its
    /// cell, so without this a declared height is overwritten by the row and
    /// then read back as if it were the item's own — lost after one pass, with
    /// the cascade still holding the right value.
    child_declared_height: std::collections::HashSet<String>,
    /// Whether a background was DECLARED for this box.
    ///
    /// A `<div>`'s initial background is `transparent`, and `colors.background`
    /// carries a light default that must never be painted — so "has a colour"
    /// cannot be the test. Without this the panel painted nothing at all and a
    /// declared `background: #2d7ff9` was stored and silently dropped: layout
    /// placed the box correctly and the window showed empty space.
    background_set: bool,
    /// A real CSS border: four widths and one colour. Distinct from
    /// [`FlowLayoutPanel::bordered`], which is the `<fieldset>` groove with a
    /// gap for its legend — a different thing that happens to also be a line.
    css_border: Edges,
    /// One colour per SIDE. CSS gives each side its own, and a single colour
    /// cannot spell `border-bottom: 1px solid #ccc` on an otherwise borderless
    /// box — the ordinary rule under a heading.
    css_border_color: [(u8, u8, u8, u8); 4],
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
    /// **The box's own text** — wxhtmledit's `Box::ownText`.
    ///
    /// A box has text AND children. That is the fact our widget set was missing
    /// and the reason `<p>a <strong>b</strong></p>` could not exist: text-bearing
    /// tags mapped to a LEAF label, a leaf refuses `add_child`, and
    /// `append_child` then put the child back in `detached` — the `<strong>`
    /// vanished with nothing reported.
    ///
    /// Two placements, because HTML has two. Bordered, it is a `<legend>`: drawn
    /// across the top edge in the gap the border leaves for it, which is what a
    /// `TGroupBox.Caption` is. Otherwise it is ordinary content, drawn inside the
    /// content box above the children — a `<p>`'s text.
    pub caption: String,
    /// The style `caption` is drawn in — same reasoning as `Label::font`, and
    /// the channel `font-family`/`font-size`/`font-weight` arrive on.
    pub font: crate::ide_text::FontSpec,
    /// `text-align` for the box's own text.
    pub text_align: crate::layout::TextAlign,
    /// **Inline children, as styled runs of this box's text** — wxhtmledit's
    /// `Box::inlineContent`.
    ///
    /// An inline element is not a box: `<strong>` inside a `<p>` has no rect
    /// and no position, it is a differently-styled stretch of the paragraph's
    /// line. So it never becomes a child widget; the DOM resolves its computed
    /// style and hands the result over here.
    ///
    /// Ordering is `caption` first, then the runs in document order. Genuine
    /// interleaving — `a <b>B</b> c` — needs TEXT NODES, which this DOM does
    /// not have: `set_text_content` replaces the whole of `caption`, so there
    /// is nowhere to put the ` c`. Stated rather than approximated.
    pub inline_content: Vec<crate::layout::InlineRun>,
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
    /// Each child's `margin` — the space it claims BETWEEN itself and its
    /// siblings, which only the container can grant.
    ///
    /// **No margin collapsing.** CSS collapses adjacent vertical margins, so
    /// 10px below one child and 10px above the next is 10px in a browser and
    /// 20px here. That is a real divergence, recorded rather than hidden;
    /// collapsing needs block formatting contexts, which this panel is not.
    child_margin: std::collections::HashMap<String, Edges>,
    /// Per-child `z-index`. Absent is `0`, and it only means anything for a
    /// **positioned** child — see [`FlowLayoutPanel::paint_order`].
    child_z: std::collections::HashMap<String, i32>,
    /// **Which formatting context this box establishes** — see [`Formatting`].
    ///
    /// `Flex` by default, so every existing caller keeps the behaviour it has:
    /// only the DOM knows an element's `display`, and only the DOM says
    /// otherwise.
    pub formatting: Formatting,
    /// Each child's **own** `display`, by child name.
    ///
    /// Normal flow has to tell block-level children from inline-level ones, and
    /// that is a fact about the CHILD which only the DOM can resolve — the
    /// widget behind a `<button>` and the one behind a `<div>` are told apart
    /// by their computed style, not by their type. Recorded by the container
    /// for the same reason `child_flex` and `child_margin` are: whoever
    /// arranges is who needs to know.
    ///
    /// Absent means block-level. A container that was never told is one no
    /// stylesheet has spoken about, and stacking is normal flow's default.
    child_display: std::collections::HashMap<String, String>,
    /// The height the last normal-flow pass actually used, including padding.
    ///
    /// A block box with `height: auto` is as tall as its content, and for a box
    /// whose content is other BOXES that number exists only once they have been
    /// placed. So the flow records it and [`Document::apply_content_height`]
    /// reads it back, rather than a second implementation guessing at what the
    /// flow was about to do.
    content_height: f32,
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
            padding: Edges::uniform(4.0),
            wrap_contents: false,
            row_gap: None,
            column_gap: None,
            reverse: false,
            align_content: "flex-start".to_string(),
            child_grid_area: std::collections::HashMap::new(),
            grid_columns: Vec::new(),
            grid_rows: Vec::new(),
            child_basis: std::collections::HashMap::new(),
            child_shrink: std::collections::HashMap::new(),
            child_declared_width: std::collections::HashSet::new(),
            child_declared_height: std::collections::HashSet::new(),
            background_set: false,
            css_border: Edges::default(),
            css_border_color: [(0, 0, 0, 255); 4],
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
            font: crate::ide_text::FontSpec::sans(14.0),
            text_align: crate::layout::TextAlign::Left,
            inline_content: Vec::new(),
            out_of_flow: std::collections::HashSet::new(),
            relative_offset: std::collections::HashMap::new(),
            child_margin: std::collections::HashMap::new(),
            child_z: std::collections::HashMap::new(),
            formatting: Formatting::Flex,
            child_display: std::collections::HashMap::new(),
            content_height: 0.0,
        }
    }

    /// The order children are PAINTED in, which is not the order they are
    /// arranged in.
    ///
    /// CSS paints in two passes: everything in normal flow first, then the
    /// **positioned** boxes on top, sorted by `z-index`. That is why a
    /// `position: relative` box with no offset still covers its non-positioned
    /// siblings, and why `z-index` does nothing at all on a static box — both
    /// are consequences of the two passes, not extra rules.
    ///
    /// A stable sort, so equal `z-index` keeps document order — which makes the
    /// default (no `z-index` anywhere) exactly the order children were added.
    fn paint_order(&self) -> Vec<usize> {
        let name = |i: usize| self.children[i].name();
        crate::layout::paint_order(
            self.children.len(),
            |i| {
                let n = name(i);
                self.out_of_flow.contains(n) || self.relative_offset.contains_key(n)
            },
            |i| self.child_z.get(name(i)).copied().unwrap_or(0),
        )
    }

    /// One child's margin, or zero.
    fn margin_of(&self, name: &str) -> Edges {
        self.child_margin.get(name).copied().unwrap_or_default()
    }

    /// The order children are LAID OUT in, honouring `row-reverse` /
    /// `column-reverse`.
    ///
    /// Reversing the sequence is exactly what reversing the main axis means:
    /// the first item is placed where the last would have been. Done here so
    /// every layout inherits it and none has to remember.
    fn flow_order(&self) -> Vec<usize> {
        let mut order = self.arrangement();
        if self.reverse {
            order.reverse();
        }
        order
    }

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
        distribute(&self.justify_content, leftover, count)
    }
}

/// Distribute free space among `count` items — the arithmetic `justify-content`
/// and `align-content` share.
///
/// One function because they ARE the same operation on different axes: one
/// spreads items along the main axis, the other spreads wrapped lines across
/// the cross axis, and CSS gives them the same keywords for that reason.
fn distribute(mode: &str, leftover: f32, count: usize) -> (f32, f32) {
    {
        let leftover = leftover.max(0.0);
        let n = count as f32;
        match mode {
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
}

impl FlowLayoutPanel {
    /// How far down the content starts.
    ///
    /// A `<legend>` occupies vertical space — content in a `<fieldset>` begins
    /// BELOW the caption, not behind it. Without this the caption drew across
    /// the first child.
    fn top_inset(&self) -> f32 {
        if self.bordered && !self.caption.is_empty() {
            self.padding.top + 14.0
        } else {
            self.padding.top
        }
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
        self.padding = Edges::uniform(padding);
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

    /// Place `widget` at `index` among the children, rather than at the end.
    ///
    /// `insertBefore` differs from `appendChild` in exactly one thing — where
    /// the child lands — so a container that took the child and ignored the
    /// index would be wrong in the only way that matters, and invisibly:
    /// the DOM would read back the right order and the window would show
    /// another. An index past the end appends, which is what an absent
    /// reference node means.
    pub fn insert(&mut self, index: usize, widget: Box<dyn PanelWidget>) {
        let rect = widget.rect();
        self.natural
            .insert(widget.name().to_string(), (rect.w, rect.h));
        let index = index.min(self.children.len());
        self.children.insert(index, widget);
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

        match self.formatting {
            Formatting::Grid => self.layout_grid(),
            Formatting::Normal => self.layout_normal_flow(),
            Formatting::Flex => match self.flow_direction {
                FlowDirection::LeftToRight => self.layout_left_to_right(),
                FlowDirection::TopDown => self.layout_top_down(),
            },
        }
    }

    /// **CSS grid** — items placed into the cells of a track template.
    ///
    /// `display: grid` parsed, cascaded and then ran normal flow, because only
    /// `flex` selected a formatting context and everything else fell through.
    /// This is the layout that was missing.
    ///
    /// Auto-placement, row-major, which is what a template alone means: items
    /// fill the columns in document order and start a new row when they run
    /// out. Explicit placement (`grid-column: 2 / 4`, `grid-area`) is **not
    /// here** — it needs a placement cursor and span resolution, and stating
    /// that is better than half-placing things.
    ///
    /// **No template is not the same as no columns.** CSS gives an untemplated
    /// grid one implicit column, so `display: grid` alone stacks its items
    /// full-width, and that is what the `1fr` fallback below produces.
    fn layout_grid(&mut self) {
        let r = self.rect;
        let inner_w = (r.w - self.padding.horizontal()).max(0.0);
        let inner_h = (r.h - self.top_inset() - self.padding.bottom).max(0.0);

        let out_of_flow = self.out_of_flow.clone();
        let items: Vec<usize> = self
            .arrangement()
            .into_iter()
            .filter(|i| !out_of_flow.contains(self.children[*i].name()))
            .collect();
        if items.is_empty() {
            self.content_height = self.padding.vertical();
            return;
        }

        let columns = if self.grid_columns.is_empty() {
            vec![crate::css::TrackSize::Fr(1.0)]
        } else {
            self.grid_columns.clone()
        };
        let col_count = columns.len();
        let (col_gap, row_gap) = (self.column_gap.unwrap_or(self.spacing), self.row_gap.unwrap_or(self.spacing));
        // **Placement first, sizing second.** Which cells an item occupies is
        // decided by the template and the item's own `grid-column`/`grid-row`,
        // never by its size — so this runs before any track is measured, and a
        // spanning item cannot change where its neighbours went.
        let areas: Vec<GridArea> = {
            let names: Vec<String> = items
                .iter()
                .map(|&i| self.children[i].name().to_string())
                .collect();
            place_grid_items(&names, &self.child_grid_area, col_count)
        };
        let row_count = areas
            .iter()
            .map(|a| a.row + a.row_span)
            .max()
            .unwrap_or(1)
            .max(1);
        let margins = self.child_margin.clone();
        let relative = self.relative_offset.clone();

        // An `auto` track is as big as the largest item in it, so the items
        // have to be measured per track BEFORE the tracks can be sized. Row
        // and column both need it; placement is known from the column count
        // alone, which is why this can happen before anything is positioned.
        let size_of = |i: usize, horizontal: bool| -> f32 {
            let name = self.children[i].name().to_string();
            let margin = margins.get(&name).copied().unwrap_or_default();
            // **The item's CURRENT size, not the one recorded when it
            // arrived.** `natural` remembers the size a FLEX pass destroyed by
            // stretching; grid stretches items to their cell, so the same trap
            // applies — and reading the recorded size pins every track to the
            // item's size at insertion, silently ignoring a `height: 64px` the
            // sheet applied afterwards. Measured: an `auto` row came out at the
            // div's 150px default with the declaration correctly computed.
            let r = self.children[i].rect();
            if horizontal {
                r.w + margin.horizontal()
            } else {
                r.h + margin.vertical()
            }
        };
        // Rows fall back to `auto` when no template names them — an implicit
        // row is content-sized, never a share of the container's height. A grid
        // taller than its box scrolls; it does not squeeze. Resolved here
        // rather than after sizing because the spanning pass below needs to
        // know which rows are `auto` and therefore allowed to grow.
        let implicit_rows: Vec<crate::css::TrackSize> = (0..row_count)
            .map(|i| {
                self.grid_rows
                    .get(i)
                    .copied()
                    .unwrap_or(crate::css::TrackSize::Auto)
            })
            .collect();
        let mut col_auto = vec![0.0f32; col_count];
        let mut row_auto = vec![0.0f32; row_count];
        // **Two passes, because a spanning item cannot be resolved against a
        // single track.** Single-track items settle their track first; only
        // then can a spanner know how much of it is already covered and
        // contribute just the shortfall — CSS Grid §12.5, "distribute extra
        // space to spanned tracks". One pass would make every track a spanner
        // touches as big as the WHOLE of it.
        for (slot, &i) in items.iter().enumerate() {
            let a = &areas[slot];
            if a.col_span == 1 {
                col_auto[a.col] = col_auto[a.col].max(size_of(i, true));
            }
            if a.row_span == 1 {
                row_auto[a.row] = row_auto[a.row].max(size_of(i, false));
            }
        }
        for (slot, &i) in items.iter().enumerate() {
            let a = &areas[slot];
            // Only `auto` tracks grow — a fixed track already has its answer
            // and a spanner must not stretch it.
            let spread = |track_auto: &mut Vec<f32>,
                              tracks: &[crate::css::TrackSize],
                              start: usize,
                              span: usize,
                              gap: f32,
                              needed: f32| {
                if span <= 1 {
                    return;
                }
                let end = (start + span).min(track_auto.len());
                let covered: f32 = track_auto[start..end].iter().sum::<f32>()
                    + gap * (span as f32 - 1.0);
                let shortfall = needed - covered;
                if shortfall <= 0.0 {
                    return;
                }
                let flexible: Vec<usize> = (start..end)
                    .filter(|t| matches!(tracks.get(*t), Some(crate::css::TrackSize::Auto)))
                    .collect();
                if flexible.is_empty() {
                    return;
                }
                let share = shortfall / flexible.len() as f32;
                for t in flexible {
                    track_auto[t] += share;
                }
            };
            spread(
                &mut col_auto,
                &columns,
                a.col,
                a.col_span,
                col_gap,
                size_of(i, true),
            );
            spread(
                &mut row_auto,
                &implicit_rows,
                a.row,
                a.row_span,
                row_gap,
                size_of(i, false),
            );
        }

        let col_sizes = resolve_tracks(&columns, inner_w, col_gap, &col_auto);
        let row_sizes = resolve_tracks(&implicit_rows, inner_h, row_gap, &row_auto);

        let origin_x = r.x + self.padding.left;
        let origin_y = r.y + self.top_inset();
        for (slot, &i) in items.iter().enumerate() {
            let a = &areas[slot];
            let (col, row) = (a.col, a.row);
            let x = origin_x
                + col_sizes[..col].iter().sum::<f32>()
                + col_gap * col as f32;
            let y = origin_y
                + row_sizes[..row].iter().sum::<f32>()
                + row_gap * row as f32;
            // A spanning item covers its tracks AND the gaps between them —
            // the gaps are inside the area, not beside it.
            let span_w = col_sizes[col..(col + a.col_span).min(col_sizes.len())]
                .iter()
                .sum::<f32>()
                + col_gap * (a.col_span as f32 - 1.0);
            let span_h = row_sizes[row..(row + a.row_span).min(row_sizes.len())]
                .iter()
                .sum::<f32>()
                + row_gap * (a.row_span as f32 - 1.0);
            let name = self.children[i].name().to_string();
            let margin = margins.get(&name).copied().unwrap_or_default();
            let (dx, dy) = relative.get(&name).copied().unwrap_or((0.0, 0.0));
            // `stretch` is the initial value of both `align-items` and
            // `justify-items`, so an item fills its cell less its margins.
            // `stretch` fills the cell, but a DECLARED size is the author's
            // answer and is kept — the same rule normal flow applies to a
            // block child's width.
            let own = self.children[i].rect();
            let w = if self.child_declared_width.contains(&name) {
                own.w
            } else {
                (span_w - margin.horizontal()).max(0.0)
            };
            let h = if self.child_declared_height.contains(&name) {
                own.h
            } else {
                (span_h - margin.vertical()).max(0.0)
            };
            self.children[i].set_rect(LayoutRect::new(
                x + margin.left + dx,
                y + margin.top + dy,
                w,
                h,
            ));
        }

        let used: f32 = row_sizes.iter().sum::<f32>() + row_gap * (row_count as f32 - 1.0);
        self.content_height = used + self.padding.vertical();
    }

    /// **CSS normal flow** — the formatting context every block container has
    /// and this engine did not.
    ///
    /// Two kinds of child and one rule each:
    ///
    /// - **Block-level** (`display: block`, and anything unrecognised): the box
    ///   fills the content width and takes a row of its own. Nothing shares a
    ///   line with it, so an open line box is closed first.
    /// - **Inline-level** (`inline`, `inline-block`): the box keeps its natural
    ///   width and sits on the current line box, which **wraps at the content
    ///   edge**. That is what makes four `<button>`s a row of four rather than
    ///   a column of four.
    ///
    /// A line box is as tall as the tallest thing on it, which is why the
    /// vertical advance happens when the line CLOSES rather than per child.
    ///
    /// **No `spacing` and no flex weights.** Both are flexbox vocabulary; CSS
    /// puts nothing between two adjacent boxes but their margins. Two buttons
    /// written with no whitespace between them are flush in a browser, and so
    /// are these — the gap in real markup comes from the whitespace text node,
    /// which is a text-node question, not a container preference.
    ///
    /// ⚠ **A declared `width` on a BLOCK-level child is overruled here.** CSS
    /// fills the content width only when the width is `auto`, and this cannot
    /// tell the two apart: the container is told each child's `display` but not
    /// whether its width was declared. `Document::fill_available_width` makes
    /// exactly that distinction for the body's children and this is the place
    /// it is missing. Inline-level children are unaffected — they keep their
    /// own width, declared or not — which is why a page styling its buttons
    /// works and a page styling a nested `<div>`'s width does not yet.
    fn layout_normal_flow(&mut self) {
        let r = self.rect;
        let content_x = r.x + self.padding.left;
        let content_top = r.y + self.top_inset();
        let content_w = (r.w - self.padding.horizontal()).max(0.0);

        let out_of_flow = self.out_of_flow.clone();
        let relative = self.relative_offset.clone();
        let margins = self.child_margin.clone();
        let displays = self.child_display.clone();
        let fixed_widths = self.child_declared_width.clone();

        // The open line box: where it starts, how far along it we are, and how
        // tall the tallest thing on it has been so far.
        let mut line_y = content_top;
        let mut cursor_x = content_x;
        let mut line_h: f32 = 0.0;

        for i in self.arrangement() {
            let name = self.children[i].name().to_string();
            if out_of_flow.contains(&name) {
                continue;
            }
            let margin = margins.get(&name).copied().unwrap_or_default();
            let (dx, dy) = relative.get(&name).copied().unwrap_or((0.0, 0.0));
            // Three cases, not two. A box we have been told is inline-level
            // keeps its width; one we have been told is block-level fills the
            // content width; and one we have **not been told about** takes a
            // row of its own but keeps its width, because filling would destroy
            // a size we cannot get back — `nat_w` is the child's live rect, so
            // a single pass under a wrong guess is permanent.
            let declared = displays.get(&name).map(String::as_str);
            let inline = matches!(declared, Some("inline") | Some("inline-block"));
            // …and a block box only fills when its width is `auto`. A declared
            // width is the author's answer and content does not overrule it —
            // the same rule `fill_available_width` applies to the body's own
            // children, applied one level down where it was missing.
            let fills = matches!(declared, Some(d) if !inline && d != "none")
                && !fixed_widths.contains(&name);
            // **The child's CURRENT size, not the one recorded when it
            // arrived.** `natural` exists because the flex pass destroys a
            // child's size by stretching it, so `align-items` needs a
            // remembered one to align within. Normal flow stretches nothing:
            // an in-flow box is either its default size or the size CSS
            // declared for it, and its rect holds whichever that is. Reading
            // the recorded size here would pin every box to its size at
            // insertion and silently ignore a `width` set afterwards — which
            // is how a page styles itself.
            let (nat_w, nat_h) = {
                let r = self.children[i].rect();
                (r.w, r.h)
            };

            if inline {
                let advance = margin.horizontal() + nat_w;
                // Wrap — but never on an EMPTY line. A box wider than its
                // container overflows it; moving it to a line of its own would
                // not make it fit and would leave a blank line above it.
                if cursor_x > content_x && cursor_x - content_x + advance > content_w {
                    line_y += line_h;
                    cursor_x = content_x;
                    line_h = 0.0;
                }
                self.children[i].set_rect(LayoutRect::new(
                    cursor_x + margin.left + dx,
                    line_y + margin.top + dy,
                    nat_w,
                    nat_h,
                ));
                cursor_x += advance;
                line_h = line_h.max(margin.vertical() + nat_h);
            } else {
                // A block box closes whatever line is open and starts below it.
                if cursor_x > content_x {
                    line_y += line_h;
                    cursor_x = content_x;
                    line_h = 0.0;
                }
                let width = if fills {
                    (content_w - margin.horizontal()).max(0.0)
                } else {
                    nat_w
                };
                self.children[i].set_rect(LayoutRect::new(
                    content_x + margin.left + dx,
                    line_y + margin.top + dy,
                    width,
                    nat_h,
                ));
                line_y += margin.vertical() + nat_h;
            }
        }
        // Close the last line, then add the bottom padding the top was already
        // charged for by `top_inset`.
        line_y += line_h;
        self.content_height = line_y - r.y + self.padding.bottom;
    }

    // Flutter Row: children side by side. Fixed-flex children keep their
    // `fixed_main` width; flex children share the leftover by weight. Full
    // height.
    fn layout_left_to_right(&mut self) {
        let r = self.rect;
        // The CONTENT box: the panel's rect less its padding. `top_inset`
        // stands in for `padding.top` because a `<fieldset>`'s legend also
        // pushes content down.
        let inner_w = (r.w - self.padding.horizontal()).max(0.0);
        let inner_h = (r.h - self.top_inset() - self.padding.bottom).max(0.0);

        let out_of_flow = self.out_of_flow.clone();
        let items: Vec<usize> = self
            .flow_order()
            .into_iter()
            // Out of flow: the child keeps the rect its own `left`/`top` gave
            // it, and takes no space from its siblings — including its margin,
            // which is space between SIBLINGS and so means nothing here.
            .filter(|i| !out_of_flow.contains(self.children[*i].name()))
            .collect();
        if items.is_empty() {
            return;
        }

        let lines = self.flex_lines(&items, inner_w);
        // **A single line keeps the container's whole cross extent**, which is
        // what `align-items: stretch` has always meant here and what every
        // Flutter `Row` depends on. Only once the content wraps does a line
        // have a height of its own to align within.
        let single = lines.len() == 1;
        let (gap, cross_gap) = (self.main_gap(), self.cross_gap());
        let relative = self.relative_offset.clone();
        let child_flex = self.child_flex.clone();
        let natural_sizes = self.natural.clone();
        let child_align = self.child_align.clone();
        let margins = self.child_margin.clone();
        let align = self.align_items.clone();
        let child_basis = self.child_basis.clone();
        let child_shrink = self.child_shrink.clone();

        // **`align-content` distributes the LINES**, and it can only run once
        // every line's height is known — which is after they have been broken
        // but before any of them is placed. A single line takes the whole cross
        // extent and has nothing to distribute, so this is inert until content
        // actually wraps.
        let line_heights: Vec<f32> = if single {
            vec![inner_h]
        } else {
            lines
                .iter()
                .map(|line| {
                    line.iter()
                        .map(|&i| {
                            let name = self.children[i].name().to_string();
                            let margin = margins.get(&name).copied().unwrap_or_default();
                            natural_sizes
                                .get(&name)
                                .map(|(_, h)| *h)
                                .unwrap_or_else(|| self.children[i].rect().h)
                                .max(1.0)
                                + margin.vertical()
                        })
                        .fold(0.0f32, f32::max)
                })
                .collect()
        };
        let used: f32 = line_heights.iter().sum::<f32>() + cross_gap * (lines.len() as f32 - 1.0);
        let (lead_cross, extra_cross) = if single {
            (0.0, 0.0)
        } else {
            distribute(&self.align_content, (inner_h - used).max(0.0), lines.len())
        };

        let mut cy = r.y + self.top_inset() + lead_cross;
        for (line_index, line) in lines.iter().enumerate() {
            let n = line.len();
            let gaps = gap * (n as f32 - 1.0);
            // Per LINE, not per container: each flex line distributes its own
            // free space, so an item's width depends on who it wrapped with.
            let mut total_flex = 0.0f32;
            let mut fixed = 0.0f32;
            let mut margin_total = 0.0f32;
            let mut line_cross = 0.0f32;
            // `flex-shrink` is weighted by the item's base size — CSS's scaled
            // shrink factor, so a big item gives up more than a small one at
            // the same factor.
            let mut shrink_scaled = 0.0f32;
            for &i in line {
                let name = self.children[i].name().to_string();
                let margin = margins.get(&name).copied().unwrap_or_default();
                margin_total += margin.horizontal();
                let f = child_flex
                    .get(&name)
                    .copied()
                    .unwrap_or_else(|| self.children[i].layout_flex());
                if f > 0.0 {
                    total_flex += f;
                }
                // **Every item contributes its BASE**, growing or not. A
                // grower's base is zero unless `flex-basis` gave it one, so
                // this is the same arithmetic as before wherever no basis is
                // declared — and the correct one where it is.
                let base = self.base_main_size(i);
                fixed += base;
                let shrink = self.child_shrink.get(&name).copied().unwrap_or(0.0);
                shrink_scaled += shrink.max(0.0) * base;
                let natural = natural_sizes
                    .get(&name)
                    .map(|(_, h)| *h)
                    .unwrap_or_else(|| self.children[i].rect().h)
                    .max(1.0);
                line_cross = line_cross.max(natural + margin.vertical());
            }
            // Margins are spoken for before anything is sized — a growing child
            // shares out what is left AFTER them, or it would grow into its own
            // sibling's margin.
            let free = inner_w - gaps - fixed - margin_total;
            let leftover = free.max(0.0);
            // Overflowing, and something declared it may shrink. Absent a
            // declaration nothing shrinks, which is this panel's long-standing
            // behaviour rather than CSS's initial value — see `child_shrink`.
            let deficit = if free < 0.0 && shrink_scaled > 0.0 {
                -free
            } else {
                0.0
            };
            // Nothing grows → the leftover is real free space, and
            // `justify-content` is what decides where it goes. With a growing
            // child there is nothing left to distribute, which is why the two
            // never fight.
            let (lead, extra) = if total_flex > 0.0 {
                (0.0, 0.0)
            } else {
                self.justify(leftover, n)
            };
            let _ = line_cross;
            let cross_extent = line_heights[line_index];

            let mut cx = r.x + self.padding.left + lead;
            for &i in line {
                let name = self.children[i].name().to_string();
                let margin = margins.get(&name).copied().unwrap_or_default();
                let child = &mut self.children[i];
                let f = child_flex
                    .get(&name)
                    .copied()
                    .unwrap_or_else(|| child.layout_flex());
                let base = if let Some(basis) = child_basis.get(&name) {
                    basis.max(0.0)
                } else if f > 0.0 {
                    0.0
                } else {
                    Self::child_fixed(child.as_ref())
                };
                let grown = if f > 0.0 && total_flex > 0.0 {
                    leftover * f / total_flex
                } else {
                    0.0
                };
                let shrunk = if deficit > 0.0 {
                    let factor = child_shrink.get(&name).copied().unwrap_or(0.0).max(0.0);
                    deficit * (factor * base) / shrink_scaled
                } else {
                    0.0
                };
                let cw = (base + grown - shrunk).max(0.0);
                let natural = natural_sizes
                    .get(&name)
                    .map(|(_, h)| *h)
                    .unwrap_or_else(|| child.rect().h)
                    .max(1.0);
                let mode = child_align.get(&name).unwrap_or(&align);
                let (offset, ch) = Self::align_with(mode, cross_extent - margin.vertical(), natural);
                // The flow slot, then the relative offset on top of it. `cx`
                // advances by the SLOT, so offsetting a child never moves its
                // siblings — the half that distinguishes relative from absolute.
                let (dx, dy) = relative.get(&name).copied().unwrap_or((0.0, 0.0));
                child.set_rect(LayoutRect::new(
                    cx + margin.left + dx,
                    cy + margin.top + offset + dy,
                    cw,
                    ch,
                ));
                cx += margin.horizontal() + cw + gap + extra;
            }
            cy += cross_extent + cross_gap + extra_cross;
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
            // ⚠ **`baseline` is approximated by `flex-start`**, and that is a
            // deliberate, stated approximation rather than the silent one it
            // replaces — it used to fall through to `stretch`, which is not
            // even the right SHAPE: a baseline-aligned item keeps its natural
            // size and a stretched one does not.
            //
            // True baseline alignment needs each child's first text baseline,
            // and `PanelWidget` exposes no such metric — a container cannot ask
            // a button where its text sits. For items of equal height the two
            // agree exactly, which is the common case (a row of labels).
            "baseline" => (0.0, natural.min(inner_cross)),
            _ => (0.0, inner_cross),
        }
    }

    // Flutter Column: children stacked top to bottom. Fixed-flex children keep
    // their `fixed_main` height; flex children share the leftover by weight.
    // Full width.
    fn layout_top_down(&mut self) {
        let r = self.rect;
        let inner_w = (r.w - self.padding.horizontal()).max(0.0);
        let inner_h = (r.h - self.top_inset() - self.padding.bottom).max(0.0);

        let out_of_flow = self.out_of_flow.clone();
        let items: Vec<usize> = self
            .flow_order()
            .into_iter()
            .filter(|i| !out_of_flow.contains(self.children[*i].name()))
            .collect();
        if items.is_empty() {
            return;
        }

        // **Column wrapping.** The mirror of the row case: `flex-direction:
        // column` with `flex-wrap: wrap` breaks into COLUMNS, and `flex_lines`
        // is already axis-agnostic because `main_gap` and `base_main_size` are
        // — the only thing that changes is which extent the line is measured
        // against.
        let lines = self.flex_lines(&items, inner_h);
        let single = lines.len() == 1;
        let (gap, cross_gap) = (self.main_gap(), self.cross_gap());
        let relative = self.relative_offset.clone();
        let child_flex = self.child_flex.clone();
        let natural_sizes = self.natural.clone();
        let child_align = self.child_align.clone();
        let margins = self.child_margin.clone();
        let align = self.align_items.clone();
        let child_basis = self.child_basis.clone();
        let child_shrink = self.child_shrink.clone();

        let line_widths: Vec<f32> = if single {
            vec![inner_w]
        } else {
            lines
                .iter()
                .map(|line| {
                    line.iter()
                        .map(|&i| {
                            let name = self.children[i].name().to_string();
                            let margin = margins.get(&name).copied().unwrap_or_default();
                            natural_sizes
                                .get(&name)
                                .map(|(w, _)| *w)
                                .unwrap_or_else(|| self.children[i].rect().w)
                                .max(1.0)
                                + margin.horizontal()
                        })
                        .fold(0.0f32, f32::max)
                })
                .collect()
        };
        let used: f32 = line_widths.iter().sum::<f32>() + cross_gap * (lines.len() as f32 - 1.0);
        let (lead_cross, extra_cross) = if single {
            (0.0, 0.0)
        } else {
            distribute(&self.align_content, (inner_w - used).max(0.0), lines.len())
        };

        let mut cx = r.x + self.padding.left + lead_cross;
        for (line_index, line) in lines.iter().enumerate() {
            let n = line.len();
            let gaps = gap * (n as f32 - 1.0);
            let mut total_flex = 0.0f32;
            let mut fixed = 0.0f32;
            let mut margin_total = 0.0f32;
            let mut shrink_scaled = 0.0f32;
            for &i in line {
                let name = self.children[i].name().to_string();
                let margin = margins.get(&name).copied().unwrap_or_default();
                margin_total += margin.vertical();
                let f = child_flex
                    .get(&name)
                    .copied()
                    .unwrap_or_else(|| self.children[i].layout_flex());
                if f > 0.0 {
                    total_flex += f;
                }
                let base = self.base_main_size(i);
                fixed += base;
                let shrink = child_shrink.get(&name).copied().unwrap_or(0.0);
                shrink_scaled += shrink.max(0.0) * base;
            }
            let free = inner_h - gaps - fixed - margin_total;
            let leftover = free.max(0.0);
            let deficit = if free < 0.0 && shrink_scaled > 0.0 {
                -free
            } else {
                0.0
            };
            let (lead, extra) = if total_flex > 0.0 {
                (0.0, 0.0)
            } else {
                self.justify(leftover, n)
            };
            let cross_extent = line_widths[line_index];

            let mut cy = r.y + self.top_inset() + lead;
            for &i in line {
                let name = self.children[i].name().to_string();
                let margin = margins.get(&name).copied().unwrap_or_default();
                let child = &mut self.children[i];
                let f = child_flex
                    .get(&name)
                    .copied()
                    .unwrap_or_else(|| child.layout_flex());
                let base = if let Some(basis) = child_basis.get(&name) {
                    basis.max(0.0)
                } else if f > 0.0 {
                    0.0
                } else {
                    Self::child_fixed(child.as_ref())
                };
                let grown = if f > 0.0 && total_flex > 0.0 {
                    leftover * f / total_flex
                } else {
                    0.0
                };
                let shrunk = if deficit > 0.0 {
                    let factor = child_shrink.get(&name).copied().unwrap_or(0.0).max(0.0);
                    deficit * (factor * base) / shrink_scaled
                } else {
                    0.0
                };
                let ch = (base + grown - shrunk).max(0.0);
                let natural = natural_sizes
                    .get(&name)
                    .map(|(w, _)| *w)
                    .unwrap_or_else(|| child.rect().w)
                    .max(1.0);
                let mode = child_align.get(&name).unwrap_or(&align);
                let (offset, cw) =
                    Self::align_with(mode, cross_extent - margin.horizontal(), natural);
                let (dx, dy) = relative.get(&name).copied().unwrap_or((0.0, 0.0));
                child.set_rect(LayoutRect::new(
                    cx + margin.left + offset + dx,
                    cy + margin.top + dy,
                    cw,
                    ch,
                ));
                cy += margin.vertical() + ch + gap + extra;
            }
            cx += cross_extent + cross_gap + extra_cross;
        }
    }

    /// The fixed main-axis size of a flex-0 child (a toolbar-height bar).
    fn child_fixed(_child: &dyn PanelWidget) -> f32 {
        44.0
    }

    /// The gap **along the main axis** — between one child and the next.
    ///
    /// Which CSS property that is depends on the direction: laying out
    /// left-to-right, the space between children is `column-gap`; stacking
    /// top-down it is `row-gap`. Falls back to `spacing` when the axis was
    /// never declared, so an unstyled panel is unchanged.
    fn main_gap(&self) -> f32 {
        match self.flow_direction {
            FlowDirection::LeftToRight => self.column_gap,
            FlowDirection::TopDown => self.row_gap,
        }
        .unwrap_or(self.spacing)
    }

    /// The gap **along the cross axis** — between one wrapped line and the next.
    /// The other property of the pair, for the same reason.
    fn cross_gap(&self) -> f32 {
        match self.flow_direction {
            FlowDirection::LeftToRight => self.row_gap,
            FlowDirection::TopDown => self.column_gap,
        }
        .unwrap_or(self.spacing)
    }

    /// **Break the in-flow children into flex lines** — CSS Flexbox §9.3.
    ///
    /// `flex-wrap` was accepted and did nothing: `wrap_contents` was set by
    /// `SetFlexWrap` and read by no layout code anywhere, so a container asking
    /// to wrap simply overflowed. This is the missing half.
    ///
    /// Lines are collected by each item's **hypothetical main size** — its base
    /// size before any growing or shrinking. A `flex-grow` item has a base of
    /// zero (`flex: 1` is `1 1 0%`), so a container of growing children still
    /// forms exactly one line however many there are, which is both what CSS
    /// says and what keeps every existing Flutter `Row` where it was.
    ///
    /// `nowrap` is not a separate path — it is one line containing everything,
    /// which is precisely what this returns when wrapping is off.
    fn flex_lines(&self, items: &[usize], available: f32) -> Vec<Vec<usize>> {
        if !self.wrap_contents {
            return vec![items.to_vec()];
        }
        let gap = self.main_gap();
        let mut lines: Vec<Vec<usize>> = Vec::new();
        let mut line: Vec<usize> = Vec::new();
        let mut used = 0.0f32;
        for &i in items {
            let name = self.children[i].name();
            let margin = self.margin_of(name);
            let base = self.base_main_size(i) + margin.horizontal();
            let advance = if line.is_empty() { base } else { gap + base };
            // Never break onto an empty line: an item wider than the container
            // overflows it, and moving it down would leave a blank line above
            // and still not make it fit.
            if !line.is_empty() && used + advance > available {
                lines.push(std::mem::take(&mut line));
                used = base;
            } else {
                used += advance;
            }
            line.push(i);
        }
        if !line.is_empty() {
            lines.push(line);
        }
        lines
    }

    /// An item's base size along the main axis, before growing or shrinking.
    ///
    /// A `flex-grow` item contributes **zero**: `flex: 1` means `flex-basis: 0`,
    /// so the item's whole size comes from the free space it is given. Only
    /// non-growing items carry a base into line-breaking.
    fn base_main_size(&self, i: usize) -> f32 {
        let child = self.children[i].as_ref();
        // A declared `flex-basis` IS the base, growing or not — that is the
        // difference between "share the space" and "share what is left after
        // my content".
        if let Some(basis) = self.child_basis.get(child.name()) {
            return basis.max(0.0);
        }
        if self.flex_of(child) > 0.0 {
            return 0.0;
        }
        Self::child_fixed(child)
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
        // **A declared background paints.** Only a declared one: the initial
        // value is `transparent`, so a box nobody has styled draws nothing and
        // its container shows through, which is what made the no-op below
        // correct for the border and wrong for the fill.
        if self.background_set {
            let ts = Transform::from_scale(scale, scale);
            let mut paint = Paint::default();
            paint.anti_alias = true;
            let (br, bg, bb, ba) = self.colors.background;
            paint.set_color_rgba8(br, bg, bb, ba);
            if let Some(rect) = Rect::from_xywh(x, y, self.rect.w, self.rect.h) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
        }
        // A declared CSS border, drawn as four filled edges so each side keeps
        // its own width — a stroked outline cannot express `border-left: 4px`
        // beside `border-top: 1px`.
        if !self.css_border.is_zero() {
            let ts = Transform::from_scale(scale, scale);
            let (w, h) = (self.rect.w, self.rect.h);
            let e = self.css_border;
            // Order is top, right, bottom, left — the order CSS names sides in,
            // and the order the colours arrive in.
            for (rect, (br, bg, bb, ba)) in [
                Rect::from_xywh(x, y, w, e.top),
                Rect::from_xywh(x + w - e.right, y, e.right, h),
                Rect::from_xywh(x, y + h - e.bottom, w, e.bottom),
                Rect::from_xywh(x, y, e.left, h),
            ]
            .into_iter()
            .zip(self.css_border_color)
            {
                let Some(rect) = rect else { continue };
                let mut paint = Paint::default();
                paint.anti_alias = true;
                paint.set_color_rgba8(br, bg, bb, ba);
                pixmap.fill_rect(rect, &paint, ts, None);
            }
        }
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

    fn insert_child(
        &mut self,
        index: usize,
        child: Box<dyn PanelWidget>,
    ) -> Option<Box<dyn PanelWidget>> {
        self.insert(index, child);
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
        // The box's own text and its inline children, as ONE line. `paint` only
        // has a pixmap; text needs the render context, so it is drawn here
        // rather than there.
        let (tr, tg, tb, ta) = self.colors.foreground;
        if self.bordered && !self.caption.is_empty() {
            // A `<legend>`: in the gap the border left across the top edge,
            // OUTSIDE the content box, which is why it ignores padding.
            ctx.draw_text(&self.caption, r.x + 12.0, r.y + 2.0, tr, tg, tb, ta);
        } else if !self.caption.is_empty() || !self.inline_content.is_empty() {
            // The inline formatting context: the box's own text first, then
            // each inline child's run, shaped TOGETHER so they share a line and
            // each advance follows the last. Drawing them separately would
            // stack every run at the same x.
            let mut spans: Vec<(String, crate::ide_text::FontSpec, cosmic_text::Color)> =
                Vec::with_capacity(self.inline_content.len() + 1);
            if !self.caption.is_empty() {
                spans.push((
                    self.caption.clone(),
                    self.font.clone(),
                    cosmic_text::Color::rgba(tr, tg, tb, ta),
                ));
            }
            for run in &self.inline_content {
                let (rr, rg, rb, ra) = run.color;
                spans.push((
                    run.text.clone(),
                    run.font.clone(),
                    cosmic_text::Color::rgba(rr, rg, rb, ra),
                ));
            }
            // Alignment needs the laid-out width, and the width needs shaping —
            // so the line is measured by drawing it off-origin first only when
            // it is not left-aligned. Left is the overwhelmingly common case
            // and pays nothing.
            let content_x = r.x + self.padding.left;
            let content_w = (r.w - self.padding.horizontal()).max(0.0);
            let y = r.y + self.padding.top;
            let x = match self.text_align {
                crate::layout::TextAlign::Left => content_x,
                _ => {
                    let advance: f32 = spans
                        .iter()
                        .map(|(text, spec, _)| {
                            crate::ide_text::measure_text_spec(
                                ctx.font_system,
                                text,
                                spec,
                                ctx.scale,
                            )
                        })
                        .sum();
                    match self.text_align {
                        crate::layout::TextAlign::Center => {
                            content_x + (content_w - advance) / 2.0
                        }
                        _ => content_x + content_w - advance,
                    }
                }
            };
            crate::ide_text::draw_rich_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                &spans,
                x.max(r.x),
                y,
                // **The line breaks at the content edge.** A block box is as
                // wide as its containing block gave it and its text wraps
                // inside; without a width the whole paragraph shaped as one
                // line and ran off the right of the box, which is also the
                // layout the box was sized against.
                Some(content_w),
                ctx.scale,
            );
        }
        for i in self.paint_order() {
            self.children[i].render(ctx);
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
            // One child's margin, addressed here because margin is space
            // between SIBLINGS and only whoever arranges them can grant it.
            WidgetCommand::Custom(name, value) if name == "SetChildMargin" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, edges)) = spec.split_once('=') {
                        if let Some(edges) = command_edges(&CommandValue::Text(edges.to_string())) {
                            if edges.is_zero() {
                                self.child_margin.remove(child);
                            } else {
                                self.child_margin.insert(child.to_string(), edges);
                            }
                            self.relayout();
                        }
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetChildZ" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, z)) = spec.rsplit_once('=') {
                        match z.trim().parse::<i32>() {
                            Ok(z) => {
                                self.child_z.insert(child.to_string(), z);
                            }
                            // `z-index: auto` is not zero-with-a-name: it means
                            // the box makes no stacking decision, which is what
                            // having no entry says.
                            Err(_) => {
                                self.child_z.remove(child);
                            }
                        }
                    }
                }
                CommandValue::None
            }
            // **Which formatting context this box establishes.** The DOM
            // resolves `display` and says; nobody else does, so a widget
            // reached by kind rather than by tag keeps flex.
            WidgetCommand::Custom(name, value) if name == "SetFormatting" => {
                if let CommandValue::Text(mode) = value {
                    let formatting = match mode.trim() {
                        "flex" => Formatting::Flex,
                        "grid" => Formatting::Grid,
                        _ => Formatting::Normal,
                    };
                    if self.formatting != formatting {
                        self.formatting = formatting;
                        self.relayout();
                    }
                }
                CommandValue::None
            }
            // One child's `display`, addressed to the container because normal
            // flow is what needs it — the difference between a box that takes a
            // row of its own and one that shares a line.
            WidgetCommand::Custom(name, value) if name == "SetChildDisplay" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, display)) = spec.rsplit_once('=') {
                        self.child_display
                            .insert(child.to_string(), display.trim().to_string());
                        self.relayout();
                    }
                }
                CommandValue::None
            }
            // **What the flow actually used.** A read, not a computation: the
            // number is whatever the last pass placed, so nothing can disagree
            // with where the children really are.
            WidgetCommand::Custom(name, _) if name == "ContentHeight" => {
                CommandValue::Number(self.content_height as f64)
            }
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
                    // The axis and its DIRECTION are two facts. Folding the
                    // reverse forms onto the forward ones kept the axis and
                    // dropped the direction.
                    self.reverse = direction.ends_with("-reverse");
                    self.relayout();
                }
                CommandValue::None
            }
            // `gap` and `padding`. Both already existed as panel fields with no
            // route in from CSS, which is the whole shape of this work: the
            // layout algorithms were here, the vocabulary could not reach them.
            // `gap` is the shorthand and sets both axes; the longhands set one
            // each. `spacing` is left alone — it is the toolkit's own
            // inter-child spacing and the fallback when an axis is undeclared,
            // so overwriting it here would make `column-gap` silently move the
            // rows as well.
            WidgetCommand::Custom(name, value) if name == "SetGap" => {
                if let Some(gap) = command_number(value) {
                    self.row_gap = Some(gap as f32);
                    self.column_gap = Some(gap as f32);
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, _) if name == "PaintsBackground" => {
                CommandValue::Bool(self.background_set)
            }
            WidgetCommand::Custom(name, _) if name == "PaintsBorder" => {
                CommandValue::Bool(!self.css_border.is_zero())
            }
            WidgetCommand::Custom(name, value) if name == "SetBorderBox" => {
                if let CommandValue::Text(spec) = value {
                    let mut parts = spec.split(';');
                    if let Some(widths) = parts.next() {
                        self.css_border =
                            command_edges(&CommandValue::Text(widths.to_string()))
                                .unwrap_or_default();
                    }
                    if let Some(colours) = parts.next() {
                        for (i, c) in colours.split(',').enumerate().take(4) {
                            if let Some(rgba) = command_color(&CommandValue::Text(c.to_string())) {
                                self.css_border_color[i] = rgba;
                            }
                        }
                    }
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value)
                if name == "SetChildWidthMode" || name == "SetChildHeightMode" =>
            {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, mode)) = spec.rsplit_once('=') {
                        let set = if name == "SetChildWidthMode" {
                            &mut self.child_declared_width
                        } else {
                            &mut self.child_declared_height
                        };
                        if mode.trim() == "declared" {
                            set.insert(child.to_string());
                        } else {
                            set.remove(child);
                        }
                        self.relayout();
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetChildFlexBasis" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, px)) = spec.rsplit_once('=') {
                        if let Ok(px) = px.trim().parse::<f32>() {
                            self.child_basis.insert(child.to_string(), px);
                            self.relayout();
                        }
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetChildFlexShrink" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, factor)) = spec.rsplit_once('=') {
                        if let Ok(factor) = factor.trim().parse::<f32>() {
                            self.child_shrink.insert(child.to_string(), factor);
                            self.relayout();
                        }
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetChildGridArea" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, area)) = spec.split_once('=') {
                        let parts: Vec<crate::css::GridLine> = area
                            .split(',')
                            .map(|p| {
                                crate::css::GridLine::parse(p)
                                    .unwrap_or(crate::css::GridLine::Auto)
                            })
                            .collect();
                        if parts.len() == 4 {
                            self.child_grid_area
                                .insert(child.to_string(), (parts[0], parts[1], parts[2], parts[3]));
                            self.relayout();
                        }
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetAlignContent" => {
                if let CommandValue::Text(mode) = value {
                    self.align_content = mode.trim().to_string();
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetGridColumns" => {
                if let CommandValue::Text(spec) = value {
                    self.grid_columns = crate::css::parse_track_list(spec).unwrap_or_default();
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetGridRows" => {
                if let CommandValue::Text(spec) = value {
                    self.grid_rows = crate::css::parse_track_list(spec).unwrap_or_default();
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetRowGap" => {
                if let Some(gap) = command_number(value) {
                    self.row_gap = Some(gap as f32);
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetColumnGap" => {
                if let Some(gap) = command_number(value) {
                    self.column_gap = Some(gap as f32);
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetPadding" => {
                if let Some(padding) = command_edges(value) {
                    self.padding = padding;
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
            // The box's own text is styled by the cascade like anyone else's —
            // see `FontSpec::apply_command`. Without this a `<p>` took a
            // container's default font and no declaration could reach it.
            WidgetCommand::Custom(name, value) if self.font.apply_command(name, value) => {
                let _ = value;
                CommandValue::None
            }
            // The DOM resolved this box's inline children into runs — see
            // `InlineRun`. Replaces wholesale rather than appending, because it
            // is a re-derivation of the same fact, not an addition to it.
            WidgetCommand::Custom(name, CommandValue::Runs(runs)) if name == "SetInlineContent" => {
                self.inline_content = runs.clone();
                CommandValue::None
            }
            WidgetCommand::Custom(name, _) if name == "GetInlineContent" => {
                CommandValue::Runs(self.inline_content.clone())
            }
            WidgetCommand::Custom(name, CommandValue::Text(value)) if name == "SetTextAlign" => {
                if let Some(align) = crate::layout::TextAlign::from_css(value) {
                    self.text_align = align;
                }
                CommandValue::None
            }
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
                    if let Some(p) = command_edges(val) {
                        self.padding = p;
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
                        self.background_set = true;
                    }
                    CommandValue::None
                }
                _ => CommandValue::None,
            },
            _ => CommandValue::None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::button::Button;

    fn panel_with_two_buttons() -> FlowLayoutPanel {
        let mut panel = FlowLayoutPanel::new();
        panel.set_rect(LayoutRect::new(0.0, 0.0, 400.0, 300.0));
        panel.add(Box::new(Button::new("a")));
        panel.add(Box::new(Button::new("b")));
        panel
    }

    /// A normal-flow panel holding two inline-level boxes.
    ///
    /// **Configured BEFORE the children arrive**, which is not a tidiness
    /// preference — `add()` arranges each child as it goes, so a panel told
    /// afterwards has already laid them out under the previous rules. Left in
    /// flex, the two buttons are stretched to the container's width first, and
    /// as inline boxes they then correctly no longer fit on one line. That is
    /// the ordering `Document::announce_child_display` exists to guarantee.
    fn inline_panel() -> FlowLayoutPanel {
        let mut panel = FlowLayoutPanel::new();
        panel.set_rect(LayoutRect::new(0.0, 0.0, 400.0, 300.0));
        panel.handle_command(&WidgetCommand::Custom(
            "SetFormatting".into(),
            CommandValue::Text("normal".into()),
        ));
        for name in ["a", "b"] {
            panel.handle_command(&WidgetCommand::Custom(
                "SetChildDisplay".into(),
                CommandValue::Text(format!("{name}=inline-block")),
            ));
        }
        panel.add(Box::new(Button::new("a")));
        panel.add(Box::new(Button::new("b")));
        panel
    }

    /// **The Flutter guard, at the level Flutter actually reaches.**
    ///
    /// `Row`/`Column` map to this widget by KIND through `vybe:gui` — no
    /// element, no tag, no cascade — so nothing in the DOM's tests pins their
    /// behaviour. The invariant is this constructor's default: a panel nobody
    /// has spoken to runs flex, stacks its children and stretches them across
    /// the cross axis, exactly as it always has.
    ///
    /// The day someone "tidies up" the default to `Normal`, every Flutter
    /// frontend silently re-lays-out and this is what says so.
    #[test]
    fn a_panel_nobody_configured_still_runs_flex() {
        let panel = panel_with_two_buttons();
        assert_eq!(panel.formatting, Formatting::Flex);

        let (a, b) = (panel.child(0).rect(), panel.child(1).rect());
        assert_ne!(a.y, b.y, "flex children must keep stacking");
        // `align-items: stretch` — the cross axis fills, which is precisely the
        // behaviour normal flow must NOT have and this must keep.
        assert!(
            a.w > 300.0,
            "a flex child stopped stretching across the cross axis: {}",
            a.w
        );
    }

    fn row_of(n: usize, flex: &str) -> FlowLayoutPanel {
        let mut panel = FlowLayoutPanel::new();
        panel.flow_direction = FlowDirection::LeftToRight;
        panel.set_rect(LayoutRect::new(0.0, 0.0, 400.0, 300.0));
        for i in 0..n {
            panel.add(Box::new(Button::new(&format!("b{i}"))));
        }
        for i in 0..n {
            panel.handle_command(&WidgetCommand::Custom(
                "SetChildFlex".into(),
                CommandValue::Text(format!("b{i}={flex}")),
            ));
        }
        panel
    }

    fn tops(panel: &FlowLayoutPanel, n: usize) -> Vec<f32> {
        (0..n).map(|i| panel.child(i).rect().y).collect()
    }

    /// **`flex-wrap` was accepted and did nothing.** `wrap_contents` was set by
    /// `SetFlexWrap` and read by NO layout code anywhere, so a container that
    /// asked to wrap simply overflowed its own edge.
    #[test]
    fn flex_wrap_breaks_a_line_and_nowrap_does_not() {
        // Non-growing children, because a growing one has a base size of zero
        // and a container of those is always exactly one line.
        let mut panel = row_of(12, "0");

        // `nowrap` is the default and CSS's: one line, however far it overflows.
        let before = tops(&panel, 12);
        assert!(
            before.iter().all(|y| *y == before[0]),
            "nowrap must keep every child on one line, got {before:?}"
        );

        panel.handle_command(&WidgetCommand::Custom(
            "SetFlexWrap".into(),
            CommandValue::Text("wrap".into()),
        ));
        let after = tops(&panel, 12);
        assert!(
            after.iter().any(|y| *y != after[0]),
            "flex-wrap: wrap did not break a line, got {after:?}"
        );
        // And a wrapped line starts back at the content edge.
        let first_x = panel.child(0).rect().x;
        let wrapped = (0..12).find(|i| panel.child(*i).rect().y != after[0]);
        let wrapped = wrapped.expect("some child wrapped");
        assert_eq!(
            panel.child(wrapped).rect().x,
            first_x,
            "a wrapped child did not return to the content edge"
        );
    }

    /// **`row-reverse` was accepted and never reversed anything** — both
    /// reverse forms folded onto their forward direction, keeping the axis and
    /// dropping the direction.
    #[test]
    fn a_reversed_direction_lays_out_backwards() {
        let mut panel = row_of(3, "0");
        let forward: Vec<f32> = (0..3).map(|i| panel.child(i).rect().x).collect();
        assert!(
            forward[0] < forward[2],
            "the forward row is not left to right"
        );

        panel.handle_command(&WidgetCommand::Custom(
            "SetFlexDirection".into(),
            CommandValue::Text("row-reverse".into()),
        ));
        let reversed: Vec<f32> = (0..3).map(|i| panel.child(i).rect().x).collect();
        assert!(
            reversed[0] > reversed[2],
            "row-reverse did not reverse the order: {reversed:?}"
        );
        // The DOCUMENT order is untouched — only where they are placed moved.
        assert_eq!(panel.child(0).name(), "b0");
    }

    /// `align-items: baseline` fell through to `stretch`, which is not even the
    /// right shape: a baseline-aligned item keeps its natural size.
    #[test]
    fn baseline_keeps_the_items_natural_size_unlike_stretch() {
        let mut panel = row_of(2, "0");
        panel.handle_command(&WidgetCommand::Custom(
            "SetAlignItems".into(),
            CommandValue::Text("stretch".into()),
        ));
        let stretched = panel.child(0).rect().h;
        panel.handle_command(&WidgetCommand::Custom(
            "SetAlignItems".into(),
            CommandValue::Text("baseline".into()),
        ));
        let aligned = panel.child(0).rect().h;
        assert!(
            aligned < stretched,
            "baseline stretched the item like `stretch` did: {aligned} vs {stretched}"
        );
    }

    /// Column wrapping — the mirror of the row case, and missing until now.
    #[test]
    fn a_column_wraps_into_a_second_column() {
        let mut panel = FlowLayoutPanel::new();
        panel.flow_direction = FlowDirection::TopDown;
        panel.set_rect(LayoutRect::new(0.0, 0.0, 400.0, 200.0));
        for i in 0..8 {
            panel.add(Box::new(Button::new(&format!("b{i}"))));
        }
        for i in 0..8 {
            panel.handle_command(&WidgetCommand::Custom(
                "SetChildFlex".into(),
                CommandValue::Text(format!("b{i}=0")),
            ));
        }
        // nowrap: one column, however far it overflows the 200px height.
        let xs: Vec<f32> = (0..8).map(|i| panel.child(i).rect().x).collect();
        assert!(
            xs.iter().all(|x| *x == xs[0]),
            "nowrap must keep one column, got {xs:?}"
        );

        panel.handle_command(&WidgetCommand::Custom(
            "SetFlexWrap".into(),
            CommandValue::Text("wrap".into()),
        ));
        let xs: Vec<f32> = (0..8).map(|i| panel.child(i).rect().x).collect();
        assert!(
            xs.iter().any(|x| *x != xs[0]),
            "a column did not wrap, got {xs:?}"
        );
    }

    /// `align-content` distributes the LINES, and is inert until they wrap.
    #[test]
    fn align_content_distributes_the_wrapped_lines() {
        let mut panel = row_of(12, "0");
        panel.handle_command(&WidgetCommand::Custom(
            "SetFlexWrap".into(),
            CommandValue::Text("wrap".into()),
        ));
        let top = panel.child(0).rect().y;

        panel.handle_command(&WidgetCommand::Custom(
            "SetAlignContent".into(),
            CommandValue::Text("center".into()),
        ));
        let centred = panel.child(0).rect().y;
        assert!(
            centred > top,
            "align-content: center did not move the first line down: {centred} vs {top}"
        );
    }

    /// `row-gap` and `column-gap` are two properties. All three spellings used
    /// to collapse onto one `spacing` scalar, so declaring one moved both axes.
    #[test]
    fn row_gap_and_column_gap_are_separate_axes() {
        let mut panel = row_of(12, "0");
        panel.handle_command(&WidgetCommand::Custom(
            "SetFlexWrap".into(),
            CommandValue::Text("wrap".into()),
        ));
        panel.handle_command(&WidgetCommand::Custom(
            "SetColumnGap".into(),
            CommandValue::Number(0.0),
        ));
        panel.handle_command(&WidgetCommand::Custom(
            "SetRowGap".into(),
            CommandValue::Number(40.0),
        ));

        // Along the row: no gap, so each child starts where the last ended.
        let (a, b) = (panel.child(0).rect(), panel.child(1).rect());
        assert_eq!(a.y, b.y, "the first two children should share a line");
        assert_eq!(
            b.x, a.x + a.w,
            "column-gap: 0 must leave no space along the main axis"
        );

        // Across the rows: 40px, which the column gap must not have touched.
        let second_line = (0..12)
            .map(|i| panel.child(i).rect())
            .find(|r| r.y != a.y)
            .expect("content wrapped to a second line");
        assert!(
            second_line.y >= a.y + a.h + 39.5,
            "row-gap did not separate the lines: line 2 at {} vs line 1 at {} high {}",
            second_line.y,
            a.y,
            a.h
        );
    }

    /// Normal flow is what a box gets when it is TOLD, and only then.
    #[test]
    fn told_to_run_normal_flow_it_lays_out_a_line() {
        let panel = inline_panel();
        let (a, b) = (panel.child(0).rect(), panel.child(1).rect());
        assert_eq!(a.y, b.y, "inline-level boxes share a line box");
        assert!(b.x >= a.x + a.w - 0.5, "the second box overlaps the first");
        assert!(a.w < 300.0, "an inline-level box was stretched: {}", a.w);
    }

    /// A margin on an inline-level box moves the box and grows the line it is
    /// on — worth pinning, because the calculator gets one the moment anyone
    /// styles a key.
    #[test]
    fn a_margin_on_an_inline_box_moves_it_and_grows_the_line() {
        let mut panel = inline_panel();
        let plain = panel.child(1).rect().x;
        let height = panel.content_height;

        // `top,right,bottom,left` as bare numbers — the wire format
        // `Document::send_child_margin` uses. A CSS length is not accepted
        // here, and passing one silently left the margin unset.
        panel.handle_command(&WidgetCommand::Custom(
            "SetChildMargin".into(),
            CommandValue::Text("a=10,10,10,10".into()),
        ));
        let a = panel.child(0).rect();
        assert!(a.x >= 10.0, "the left margin did not move the box: {}", a.x);
        assert!(
            panel.child(1).rect().x >= plain + 20.0,
            "the margin claimed no space from the sibling beside it"
        );
        assert!(
            panel.content_height > height,
            "a vertical margin on an inline-block did not grow its line box"
        );
    }
}

/// **Resolve a grid track list to pixel sizes.**
///
/// The order is the whole algorithm and it is not negotiable: fixed tracks take
/// what they asked for, `auto` tracks take what their contents need, and only
/// then is what remains divided among the `fr` tracks by weight. An `fr` track
/// is a share of the LEFTOVER, so it cannot be sized until everything that is
/// not leftover has been subtracted — including the gaps, which are space the
/// grid spends before any track sees it.
///
/// With no `fr` track the remainder simply goes unused, which is what CSS does:
/// a grid of fixed columns does not stretch to fill its container.
fn resolve_tracks(
    tracks: &[crate::css::TrackSize],
    extent: f32,
    gap: f32,
    auto_sizes: &[f32],
) -> Vec<f32> {
    use crate::css::TrackSize;
    let gaps = gap * (tracks.len() as f32 - 1.0).max(0.0);
    let mut sizes = vec![0.0f32; tracks.len()];
    let mut fr_total = 0.0f32;
    let mut used = gaps;
    for (i, track) in tracks.iter().enumerate() {
        match *track {
            TrackSize::Px(v) => sizes[i] = v.max(0.0),
            TrackSize::Percent(p) => sizes[i] = (extent * p / 100.0).max(0.0),
            TrackSize::Auto => sizes[i] = auto_sizes.get(i).copied().unwrap_or(0.0).max(0.0),
            TrackSize::Fr(f) => {
                fr_total += f.max(0.0);
                continue;
            }
        }
        used += sizes[i];
    }
    if fr_total > 0.0 {
        let leftover = (extent - used).max(0.0);
        for (i, track) in tracks.iter().enumerate() {
            if let TrackSize::Fr(f) = *track {
                sizes[i] = leftover * f.max(0.0) / fr_total;
            }
        }
    }
    sizes
}

/// Where one item sits on the grid, in 0-based track coordinates.
#[derive(Clone, Copy, Debug)]
struct GridArea {
    col: usize,
    row: usize,
    col_span: usize,
    row_span: usize,
}

/// **Grid placement** — CSS Grid §8, the auto-placement algorithm in the shape
/// this engine needs.
///
/// An item is either *explicitly placed* (it named a line) or *auto-placed*
/// (the cursor finds it a home). Explicit items are placed FIRST and occupy
/// their cells, so an auto-placed item flows around them rather than through
/// them — which is the whole point of being able to pin one item.
///
/// Auto-placement is row-major (`grid-auto-flow: row`, the initial value) and
/// **sparse**: the cursor never moves backwards, so an item that does not fit
/// the rest of the current row starts a new one rather than back-filling an
/// earlier hole. That is what CSS specifies for the default flow.
fn place_grid_items(
    names: &[String],
    declared: &std::collections::HashMap<
        String,
        (
            crate::css::GridLine,
            crate::css::GridLine,
            crate::css::GridLine,
            crate::css::GridLine,
        ),
    >,
    col_count: usize,
) -> Vec<GridArea> {
    use crate::css::GridLine;
    let cols = col_count.max(1);

    // How many tracks an item covers, from whichever pair of lines it gave.
    // A start and an end line span the distance between them; a `span` says the
    // distance directly; anything else is one track.
    let span_of = |start: GridLine, end: GridLine| -> usize {
        match (start, end) {
            (_, GridLine::Span(n)) => (n as usize).max(1),
            (GridLine::Span(n), _) => (n as usize).max(1),
            (GridLine::Line(a), GridLine::Line(b)) => (b - a).unsigned_abs() as usize,
            _ => 1,
        }
        .max(1)
    };
    // A 1-based CSS grid line to a 0-based track index.
    let track_of = |line: GridLine| -> Option<usize> {
        match line {
            GridLine::Line(n) if n >= 1 => Some((n - 1) as usize),
            // A negative line counts back from the END of the explicit grid,
            // which is what `grid-column: -1` means: the last line.
            GridLine::Line(n) if n <= -1 => {
                let from_end = cols as i32 + 1 + n;
                (from_end >= 1).then(|| (from_end - 1) as usize)
            }
            _ => None,
        }
    };

    let mut occupied: Vec<Vec<bool>> = Vec::new();
    let fits = |col: usize, row: usize, cs: usize, rs: usize, grid: &mut Vec<Vec<bool>>| {
        while grid.len() < row + rs {
            grid.push(vec![false; cols]);
        }
        (row..row + rs).all(|r| (col..(col + cs).min(cols)).all(|c| !grid[r][c]))
    };
    let occupy = |col: usize, row: usize, cs: usize, rs: usize, grid: &mut Vec<Vec<bool>>| {
        while grid.len() < row + rs {
            grid.push(vec![false; cols]);
        }
        for r in row..row + rs {
            for c in col..(col + cs).min(cols) {
                grid[r][c] = true;
            }
        }
    };

    let mut areas: Vec<Option<GridArea>> = vec![None; names.len()];

    // Pass 1: everything that named a column line. Explicit first, so the
    // cursor in pass 2 sees them as taken.
    for (i, name) in names.iter().enumerate() {
        let Some(&(cs, ce, rs, re)) = declared.get(name) else {
            continue;
        };
        let (Some(col), Some(row)) = (track_of(cs), track_of(rs)) else {
            continue;
        };
        let (col_span, row_span) = (span_of(cs, ce), span_of(rs, re));
        let col = col.min(cols.saturating_sub(1));
        occupy(col, row, col_span, row_span, &mut occupied);
        areas[i] = Some(GridArea {
            col,
            row,
            col_span,
            row_span,
        });
    }

    // Pass 2: everything else, in document order, sparsely.
    let (mut cursor_col, mut cursor_row) = (0usize, 0usize);
    for (i, name) in names.iter().enumerate() {
        if areas[i].is_some() {
            continue;
        }
        let (cs, ce, rs, re) = declared
            .get(name)
            .copied()
            .unwrap_or((GridLine::Auto, GridLine::Auto, GridLine::Auto, GridLine::Auto));
        let (col_span, row_span) = (span_of(cs, ce).min(cols), span_of(rs, re));
        // A column-pinned item keeps its column and only looks for a free row.
        let pinned = track_of(cs).map(|c| c.min(cols.saturating_sub(1)));
        loop {
            let col = pinned.unwrap_or(cursor_col);
            if col + col_span > cols {
                cursor_col = 0;
                cursor_row += 1;
                continue;
            }
            if fits(col, cursor_row, col_span, row_span, &mut occupied) {
                occupy(col, cursor_row, col_span, row_span, &mut occupied);
                areas[i] = Some(GridArea {
                    col,
                    row: cursor_row,
                    col_span,
                    row_span,
                });
                if pinned.is_none() {
                    cursor_col = col + col_span;
                    if cursor_col >= cols {
                        cursor_col = 0;
                        cursor_row += 1;
                    }
                }
                break;
            }
            cursor_col += 1;
            if cursor_col >= cols {
                cursor_col = 0;
                cursor_row += 1;
            }
        }
    }

    areas
        .into_iter()
        .map(|a| {
            a.unwrap_or(GridArea {
                col: 0,
                row: 0,
                col_span: 1,
                row_span: 1,
            })
        })
        .collect()
}
