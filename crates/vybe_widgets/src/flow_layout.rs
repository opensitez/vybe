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
    /// `Row`/`Column` mean — they reach this widget by KIND
    /// and never become elements, so this stays their behaviour exactly.
    Flex,
    /// **CSS normal flow** — block-level children stack, inline-level children
    /// share line boxes and wrap at the content edge.
    Normal,
    /// `display: grid`. Items are placed into the cells of a track template.
    Grid,
    /// **`display: table`** — CSS 2.1 §17.
    ///
    /// Not a grid with different names. A grid's tracks come from a template
    /// the author wrote; a table's COLUMNS are measured from what the cells
    /// contain, and a cell may span several of them in either axis. The table
    /// box also reaches through boxes that establish nothing of their own —
    /// rows and row groups — because a column exists ACROSS rows that know
    /// nothing about each other.
    Table,
    /// **A row establishes no formatting context** — CSS 2.1 §17.5.
    ///
    /// A `<tr>` holds cells in the tree and positions none of them: a column
    /// spans rows that know nothing about each other, so only the table can
    /// place a cell. Without this the row ran normal flow over its own
    /// children and stacked the cells the table had just laid out side by side,
    /// each stretched to the row's width — the row's layout ran last and won.
    ///
    /// Marking each cell out-of-flow was the first attempt and does not hold:
    /// the DOM re-asserts `flow` placement for every child that is not
    /// positioned, so the flag is erased on the next restyle. Doing nothing at
    /// ALL is not a flag on a child; it is what a row IS.
    TableRow,
}

/// **One thing sitting on a line box** — a word, or an atomic inline-level box.
///
/// The inline formatting context's output, and the reason it has one. Placement
/// used to be computed TWICE and independently: `layout_normal_flow` advanced a
/// scalar cursor to position the widgets, and `render` advanced a second one to
/// draw the text. Both were right only while nothing wrapped, because neither
/// measured with a wrap width — so a paragraph in a narrow column painted its
/// box at the start of the wrong line, on top of the text it should follow, and
/// shaped everything after that box into a column one character wide.
///
/// Deciding it once and drawing what was decided is what makes those two
/// answers impossible to disagree. `render` reads this; it computes nothing.
#[derive(Clone, Debug)]
struct InlineItem {
    /// The word, with its trailing spaces. Empty for an atomic slot.
    text: String,
    font: crate::ide_text::FontSpec,
    color: (u8, u8, u8, u8),
    /// The child widget occupying this slot — see [`crate::layout::InlineRun`].
    /// The box draws itself, so this item only reserves the room.
    atomic: Option<String>,
    /// Which element the word came from, for the hit test.
    source: Option<String>,
    cursor: Option<crate::css::Cursor>,
    /// Relative to the content origin, so a box that MOVES does not invalidate
    /// its own inline layout.
    x: f32,
    y: f32,
    w: f32,
    h: f32,
}

/// Break `text` into the pieces a line can break BETWEEN.
///
/// A word carries its own trailing spaces: the break happens between words, so
/// the space that separates two of them belongs to the one before it. Splitting
/// them apart would let a line start with a space, which is the one thing CSS
/// white-space processing is most insistent it must not do.
fn split_words(text: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut start = 0usize;
    let mut spacing = false;
    for (at, ch) in text.char_indices() {
        if ch.is_whitespace() {
            spacing = true;
        } else if spacing {
            out.push(&text[start..at]);
            start = at;
            spacing = false;
        }
    }
    if start < text.len() {
        out.push(&text[start..]);
    }
    out
}

/// Where a row IS in the tree — a table's child, or a row group's.
///
/// **A row is not always the table's own child**, which is the whole reason
/// this is a path and not an index. `<tr>` may sit directly under `<table>`,
/// or under a `<thead>`/`<tbody>`/`<tfoot>` — and the table has to reach both
/// the same way, because a column spans rows that may be in different groups.
///
/// Addressing a row by its index among the table's children treated a whole
/// `<tbody>` AS a row: `cells_of` then looked for cells among its children,
/// found `<tr>`s instead, and every grouped table rendered completely empty.
/// The [`RowPath::row`] of the ANONYMOUS row — CSS 2.1 §17.2.1.
///
/// A cell that is not inside a row still belongs to one: the table generates an
/// anonymous row box around it. That is not a convenience, it is what makes
/// `<table><th>a</th><th>b</th></table>` a header rather than nothing, and it
/// is how a toolkit that appends CELLS to a grid gets a row without ever
/// saying the word.
const ANONYMOUS_ROW: usize = usize::MAX;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct RowPath {
    /// The row group holding this row, if any — an index into the table's
    /// children. `None` for a `<tr>` written directly under the table.
    group: Option<usize>,
    /// Index into whichever box holds the row: the group's children when
    /// grouped, the table's own otherwise.
    row: usize,
}

/// One cell's place in the table — where it lives in the tree, and which slots
/// it occupies in the grid.
///
/// The two are genuinely different questions. `row`/`index` say where the cell
/// IS: a child of a row, because that is what the markup nests. `origin` and
/// the spans say where it SITS: a column is not a sibling index once anything
/// above it spans rows.
#[derive(Clone, Copy, Debug)]
struct TableCell {
    /// The ROW holding this cell — a path, because it may be inside a group.
    row: RowPath,
    /// Index into that row's children.
    index: usize,
    colspan: u32,
    rowspan: u32,
    /// The (row, column) slot this cell starts at. Every other slot it covers
    /// holds a copy pointing back here, which is how "is this the cell's own
    /// slot" is answered without a second list.
    origin: (usize, usize),
}

/// **The inline formatting context** — CSS 2.1 §9.4.2.
///
/// Fills line boxes with words and atomic inline-level boxes, breaking to a new
/// line when what comes next no longer fits. One writer serves both readers:
/// layout asks it where each inline child goes, paint asks it where each word
/// went, and neither computes a position of its own. That is the whole fix —
/// the two used to advance separate cursors, agreed only while nothing wrapped,
/// and diverged completely the moment a line broke.
struct LineWriter {
    /// The content width lines break against. **Per line rather than per
    /// block** is what a float would need — `float` shortens only the lines it
    /// vertically overlaps — so this is the value that becomes a rectangle
    /// query when floats land. It is a single width today and stated as such.
    width: f32,
    nowrap: bool,
    align: crate::layout::TextAlign,
    /// How far along the open line we are, from the content left.
    x: f32,
    /// The open line's top, from the content top.
    y: f32,
    /// The tallest thing on the open line so far.
    height: f32,
    /// Where the open line's items begin, so alignment can shift them.
    line_start: usize,
    items: Vec<InlineItem>,
}

impl LineWriter {
    fn new(width: f32, nowrap: bool, align: crate::layout::TextAlign) -> Self {
        Self {
            width,
            nowrap,
            align,
            x: 0.0,
            y: 0.0,
            height: 0.0,
            line_start: 0,
            items: Vec::new(),
        }
    }

    /// Close the open line and start one below it.
    fn break_line(&mut self) {
        self.align_line();
        self.y += self.height;
        self.x = 0.0;
        self.height = 0.0;
        self.line_start = self.items.len();
    }

    /// `text-align` for the line just closed.
    ///
    /// **A line holding an atomic box is left-aligned**, and that is a real
    /// limitation rather than an oversight: the box's position has already been
    /// handed to the child widget by the time the line closes, so shifting the
    /// items here would move the words and leave the widget behind. Aligning
    /// both needs the child rects written after the line is complete.
    fn align_line(&mut self) {
        let free = (self.width - self.x).max(0.0);
        let shift = match self.align {
            crate::layout::TextAlign::Left => 0.0,
            crate::layout::TextAlign::Center => free / 2.0,
            _ => free,
        };
        if shift <= 0.0 || self.items[self.line_start..].iter().any(|i| i.atomic.is_some()) {
            return;
        }
        for item in &mut self.items[self.line_start..] {
            item.x += shift;
        }
    }

    /// Place one run's words, breaking where they stop fitting.
    ///
    /// Measuring the whole run instead — which is what this did, with no wrap
    /// width at all — yields a single advance wider than the box, so the cursor
    /// walks off the content edge and never returns. Every position computed
    /// after it is then wrong, which is why one missing argument put a widget on
    /// the wrong line AND shaped the text after it into a one-character column.
    fn text_run(&mut self, run: &crate::layout::InlineRun) {
        for word in split_words(&run.text) {
            let (w, h) = crate::ide_text::measure_rich_text(
                &[(
                    word.to_string(),
                    run.font.clone(),
                    cosmic_text::Color::rgb(0, 0, 0),
                )],
                None,
            );
            // Never break on an EMPTY line: a word wider than the box overflows
            // it, and moving it down would not make it fit. Never break before
            // whitespace either — trailing spaces hang past the content edge
            // rather than pushing themselves onto the next line.
            if !self.nowrap
                && self.x > 0.0
                && self.x + w > self.width
                && !word.trim_start().is_empty()
            {
                self.break_line();
            }
            self.items.push(InlineItem {
                text: word.to_string(),
                font: run.font.clone(),
                color: run.color,
                atomic: None,
                source: run.source.clone(),
                cursor: run.cursor,
                x: self.x,
                y: self.y,
                w,
                h,
            });
            self.x += w;
            self.height = self.height.max(h);
        }
    }

    /// Place an atomic inline-level box and answer where it goes, relative to
    /// the content origin. `w` is the MARGIN box — the room the line gives up.
    fn atomic(&mut self, name: &str, w: f32, h: f32) -> (f32, f32) {
        if self.x > 0.0 && self.x + w > self.width {
            self.break_line();
        }
        let at = (self.x, self.y);
        self.items.push(InlineItem {
            text: String::new(),
            font: crate::ide_text::FontSpec::sans(0.0),
            color: (0, 0, 0, 0),
            atomic: Some(name.to_string()),
            source: None,
            cursor: None,
            x: self.x,
            y: self.y,
            w,
            h,
        });
        self.x += w;
        self.height = self.height.max(h);
        at
    }

    /// A block-level box closes the open line and takes a band of its own.
    fn block(&mut self, h: f32) -> f32 {
        if self.x > 0.0 {
            self.break_line();
        }
        let top = self.y;
        self.y += h;
        top
    }

    /// Close the last line. Answers the height of everything placed.
    fn finish(&mut self) -> f32 {
        self.align_line();
        self.y + self.height
    }
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
    grid_columns: crate::css::TrackTemplate,
    grid_rows: crate::css::TrackTemplate,
    /// `grid-template-areas`, one row of cell names per entry.
    grid_areas: Vec<Vec<String>>,
    /// Per-child `grid-area: <name>` — which named area an item claims.
    child_grid_name: std::collections::HashMap<String, String>,
    /// `grid-auto-rows`/`grid-auto-columns` — the size of a track the template
    /// never named. `Auto` unless declared, which is the CSS initial value.
    grid_auto_rows: crate::css::TrackSize,
    grid_auto_columns: crate::css::TrackSize,
    /// `grid-auto-flow`, as `(fill columns first?, dense?)`.
    grid_flow_column: bool,
    grid_flow_dense: bool,
    /// `justify-items` and per-child `justify-self` — the inline axis.
    justify_items: String,
    child_justify_self: std::collections::HashMap<String, String>,
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
    /// Whether THIS box's own `height` was declared — the same fact as
    /// `child_declared_height`, asked at the other end.
    ///
    /// It decides whether the main size is DEFINITE, which flexbox §9.2.3 and
    /// §9.7 both turn on. A `flex: 1` item's basis is `0%`, and a percentage of
    /// an indefinite size resolves to `content`, not to zero — so a column with
    /// `height: auto` sizes its items to their content and then sizes itself to
    /// their sum, instead of dividing a height it does not have yet.
    ///
    /// Without it a Flutter `Column` of four `Expanded` rows divided
    /// `default_size`'s 150px guess four ways, gave each row ~37px, and drew
    /// four rows of buttons overlapping at half height. Growing could not fix
    /// it either: the content height was measured FROM the guess, so it was a
    /// fixed point that reproduced 150 for ever.
    ///
    /// Only the block axis. A block-level box's inline size is always definite
    /// — `width: auto` fills the containing block, and a stretched flex item's
    /// cross size is definite too — so a row keeps dividing its width.
    declared_height: bool,
    /// `white-space: nowrap` (or `pre`) — the box's text stays on ONE line and
    /// overflows rather than breaking.
    nowrap: bool,
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
    /// Where each named run was actually PAINTED, so a click can find it.
    ///
    /// A run is not a widget: it has no rect of its own and nothing in the
    /// widget tree to hit-test, which is why an `<a>` that becomes a run stops
    /// being clickable unless something records where it landed. This is that
    /// record, and it is written by `render` rather than by layout because the
    /// x of a run is a SHAPING result — it depends on every run before it on
    /// the line, and only the text pass knows those advances.
    ///
    /// Empty until the first paint. A click before anything has been drawn has
    /// no geometry to consult and correctly finds nothing.
    run_rects: Vec<(String, LayoutRect, Option<crate::css::Cursor>)>,
    /// The laid-out line boxes — see [`InlineItem`]. Written by layout, read by
    /// paint and by the hit test.
    inline_items: Vec<InlineItem>,
    /// `colspan` — how many COLUMNS this box occupies when it is a table cell.
    ///
    /// On the cell rather than on the table because that is where the attribute
    /// is written, and read back by the table when it builds its grid. One, not
    /// zero, for every box that is not a cell — see the `colspan` arm in
    /// `Document::set_attribute` for why the floor matters.
    colspan: u32,
    /// `rowspan` — the same fact on the block axis.
    rowspan: u32,
    /// `border-spacing`, in logical pixels — the gap the SEPARATE border model
    /// leaves between cell borders, and around the outside of the table.
    border_spacing: f32,
    /// `border-collapse: collapse`. Adjacent borders become one and the spacing
    /// goes to zero.
    border_collapse: bool,
    /// `table-layout: fixed` — take the column widths from the columns and the
    /// FIRST ROW alone and never measure the rest.
    table_layout_fixed: bool,
    /// Clicks landed on a run, waiting to be drained. A panel had none of its
    /// own before — every event came from a child widget — because a box with
    /// only text had nothing clickable in it. A run IS clickable, so the box
    /// that paints it is the thing that reports it.
    pending_events: Vec<WidgetEvent>,
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
            grid_columns: crate::css::TrackTemplate::default(),
            grid_rows: crate::css::TrackTemplate::default(),
            grid_areas: Vec::new(),
            child_grid_name: std::collections::HashMap::new(),
            grid_auto_rows: crate::css::TrackSize::Auto,
            grid_auto_columns: crate::css::TrackSize::Auto,
            grid_flow_column: false,
            grid_flow_dense: false,
            justify_items: "stretch".to_string(),
            child_justify_self: std::collections::HashMap::new(),
            child_basis: std::collections::HashMap::new(),
            child_shrink: std::collections::HashMap::new(),
            child_declared_width: std::collections::HashSet::new(),
            child_declared_height: std::collections::HashSet::new(),
            // **Definite until the cascade says otherwise.** A panel built
            // directly by the widget layer — a WinForms form, a VB designer
            // panel — was given a real rect by whoever built it, and dividing
            // it among flex weights is what it has always done. Only a CSS box
            // can be `height: auto`, and only the DOM knows, so only the DOM
            // flips this.
            nowrap: false,
            declared_height: true,
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
            run_rects: Vec::new(),
            inline_items: Vec::new(),
            colspan: 1,
            rowspan: 1,
            // §17.6.1's initial value, and the reason an unstyled HTML table has
            // visible gaps between its cells rather than a solid block.
            border_spacing: 2.0,
            border_collapse: false,
            table_layout_fixed: false,
            pending_events: Vec::new(),
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
            Formatting::Grid => {
                self.layout_grid();
                self.layout_inline_text();
            }
            // A table's own text is its caption, which is a BOX (`<caption>`
            // is `display: table-caption`), not a run — so there is no inline
            // pass here.
            Formatting::Table => self.layout_table(),
            // Deliberately nothing — see [`Formatting::TableRow`]. The row's own
            // rect is set by its table; its cells are placed by the same pass.
            Formatting::TableRow => {}
            // Normal flow lays out the text and the children TOGETHER — they
            // share line boxes — so it writes `inline_items` itself.
            Formatting::Normal => self.layout_normal_flow(),
            Formatting::Flex => {
                match self.flow_direction {
                    FlowDirection::LeftToRight => self.layout_left_to_right(),
                    FlowDirection::TopDown => self.layout_top_down(),
                }
                self.layout_inline_text();
            }
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

        // **The area template defines the grid's shape when no track list
        // does.** `grid-template-areas` with two names per row IS a two-column
        // grid; requiring `grid-template-columns` as well would make the
        // commonest spelling of a named layout silently one column.
        // **Auto-repeat resolves HERE**, where the extent is known — that is
        // the whole reason the template is carried rather than expanded.
        let columns = if !self.grid_columns.is_empty() {
            self.grid_columns.resolve(inner_w, self.column_gap.unwrap_or(self.spacing))
        } else if let Some(row) = self.grid_areas.first() {
            vec![crate::css::TrackSize::Fr(1.0); row.len()]
        } else {
            vec![crate::css::TrackSize::Fr(1.0)]
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
            place_grid_items(
                &names,
                &self.child_grid_area,
                &self.child_grid_name,
                &self.grid_areas,
                col_count,
                self.grid_flow_dense,
            )
        };
        let row_count = areas
            .iter()
            .map(|a| a.row + a.row_span)
            .max()
            .unwrap_or(1)
            .max(self.grid_areas.len())
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
        // An IMPLICIT row — one the template never named — is sized by
        // `grid-auto-rows`, not by `auto`. That property existed nowhere, so a
        // grid with more items than cells always grew content-sized rows.
        let declared_rows = self.grid_rows.resolve(inner_h, row_gap);
        let implicit_rows: Vec<crate::css::TrackSize> = (0..row_count)
            .map(|i| {
                declared_rows
                    .get(i)
                    .copied()
                    .filter(|_| !self.grid_rows.is_empty())
                    .unwrap_or(self.grid_auto_rows)
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
            // **Both axes align independently in a grid**, which is the
            // difference from flex: `align-*` places the item in its cell down
            // the block axis, `justify-*` across the inline one. `stretch` (the
            // initial value of both) fills the cell, and a DECLARED size wins
            // over either — the author's answer is not an alignment.
            let justify = self
                .child_justify_self
                .get(&name)
                .unwrap_or(&self.justify_items)
                .clone();
            let align = self.child_align.get(&name).unwrap_or(&self.align_items).clone();
            let cell_w = (span_w - margin.horizontal()).max(0.0);
            let cell_h = (span_h - margin.vertical()).max(0.0);
            let (jx, w) = if self.child_declared_width.contains(&name) {
                (Self::align_with(&justify, cell_w, own.w).0, own.w)
            } else {
                let (off, size) = Self::align_with(&justify, cell_w, own.w.max(1.0));
                (off, size)
            };
            let (ay, h) = if self.child_declared_height.contains(&name) {
                (Self::align_with(&align, cell_h, own.h).0, own.h)
            } else {
                let (off, size) = Self::align_with(&align, cell_h, own.h.max(1.0));
                (off, size)
            };
            let (x, y) = (x + jx, y + ay);
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

        // The open line box — see [`LineWriter`]. It holds the line state that
        // three separate scalars used to, and it is the same object paint reads
        // back, so the text and the widgets on a line cannot be placed by two
        // mechanisms that have never heard of each other.
        let mut line = LineWriter::new(content_w, self.nowrap, self.text_align);
        // **How far into the box's own text the line has got.** `inline_content`
        // is the sequence in DOCUMENT order — text runs and atomic slots for
        // the inline-level children — so walking it alongside `arrangement()`
        // is what puts a `<label>`'s text and the `<input>` after it on one
        // line. Before this, the text was painted from the content origin and
        // the widgets were laid out from the same origin, so a label rendered
        // on top of the field it labels.
        let runs = self.inline_content.clone();
        let mut run_cursor = 0usize;
        // The box's OWN characters lead its line: a `<p>`'s text comes before
        // anything nested inside it. A `<fieldset>`'s caption is its legend and
        // is drawn in the border gap instead, so it takes no part here.
        let caption = self.caption.clone();
        if !caption.is_empty() && !self.bordered {
            line.text_run(&crate::layout::InlineRun {
                text: caption,
                font: self.font.clone(),
                color: self.colors.foreground,
                source: None,
                atomic: None,
                cursor: None,
            });
        }

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
            // `inline-flex`/`inline-grid` are inline-level to their PARENT —
            // that is the entire meaning of the prefix, and the reason they
            // belong here rather than beside the formatting contexts.
            let inline = matches!(
                declared,
                Some("inline") | Some("inline-block") | Some("inline-flex") | Some("inline-grid")
            );
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
                // The text that comes BEFORE this box on the line. Consumed
                // here rather than measured up front, because which text
                // precedes which box is the sequence's answer, not a guess —
                // and it is the only reason the two can share a line at all.
                while let Some(run) = runs.get(run_cursor) {
                    match run.atomic.as_deref() {
                        // This box's own slot: the walk has caught up.
                        Some(slot) if slot == name => {
                            run_cursor += 1;
                            break;
                        }
                        // A DIFFERENT box's slot — that child is placed by its
                        // own turn in `arrangement()`, so stop rather than
                        // consume it here.
                        Some(_) => break,
                        None => {
                            line.text_run(run);
                            run_cursor += 1;
                        }
                    }
                }
                let (x, y) = line.atomic(
                    &name,
                    margin.horizontal() + nat_w,
                    margin.vertical() + nat_h,
                );
                self.children[i].set_rect(LayoutRect::new(
                    content_x + x + margin.left + dx,
                    content_top + y + margin.top + dy,
                    nat_w,
                    nat_h,
                ));
            } else {
                // A block box closes whatever line is open and starts below it.
                let top = line.block(margin.vertical() + nat_h);
                let width = if fills {
                    (content_w - margin.horizontal()).max(0.0)
                } else {
                    nat_w
                };
                self.children[i].set_rect(LayoutRect::new(
                    content_x + margin.left + dx,
                    content_top + top + margin.top + dy,
                    width,
                    nat_h,
                ));
            }
        }
        // **The text after the last inline child.** The walk above consumes
        // runs only when it reaches a box, so everything following the final
        // one was never placed at all — the tail of every sentence that ended
        // in words rather than in a widget.
        while let Some(run) = runs.get(run_cursor) {
            if run.atomic.is_none() {
                line.text_run(run);
            }
            run_cursor += 1;
        }
        // Close the last line, then add the bottom padding the top was already
        // charged for by `top_inset`.
        let height = line.finish();
        self.inline_items = std::mem::take(&mut line.items);
        self.content_height = (content_top - r.y) + height + self.padding.bottom;
    }

    /// **The table formatting context** — CSS 2.1 §17.5, ported from
    /// `wxhtmledit/src/TableLayout.cpp`.
    ///
    /// Four passes, in this order, because each needs the one before it:
    ///
    /// 1. **Collect the rows**, reaching through `thead`/`tbody`/`tfoot` — the
    ///    row groups establish no formatting context of their own, they only
    ///    say where their rows go in visual order (header, body, footer, no
    ///    matter what order the markup put them in).
    /// 2. **Build the cell grid**, so `rowspan` from a row above occupies the
    ///    slot a later row would otherwise fill. Without the grid, spans are
    ///    guesswork: a cell's column is not its index among its siblings.
    /// 3. **Size the columns.** `fixed` reads the first row and stops; `auto`
    ///    measures every cell. This is the pass that makes a table a table
    ///    rather than a grid — nobody declared these widths.
    /// 4. **Lay the cells out and take the row heights from them**, then place
    ///    each cell at its column's x and its row's y.
    ///
    /// The ROW boxes are given the table's width and their row's height, so a
    /// row can still paint a background; the cells are positioned in the
    /// table's own coordinates rather than the row's, because a spanning cell
    /// belongs to several rows and can be a child of only one.
    fn layout_table(&mut self) {
        let r = self.rect;
        let content_x = r.x + self.padding.left;
        let content_top = r.y + self.top_inset();
        let content_w = (r.w - self.padding.horizontal()).max(0.0);
        // §17.6.2: collapsing the borders removes the spacing between cells.
        let spacing = if self.border_collapse {
            0.0
        } else {
            self.border_spacing
        };

        let rows = self.table_rows();
        if rows.is_empty() {
            self.content_height = (content_top - r.y) + self.padding.bottom;
            return;
        }

        let grid = self.build_cell_grid(&rows);
        let columns = grid.first().map(|row| row.len()).unwrap_or(0);
        if columns == 0 {
            self.content_height = (content_top - r.y) + self.padding.bottom;
            return;
        }

        // The room the cells themselves get: the table less the spacing that
        // runs between them and around the outside.
        let cell_area = (content_w - spacing * (columns as f32 + 1.0)).max(0.0);
        let widths = self.column_widths(&grid, columns, cell_area);

        // ── Lay each cell out at its column width, and take the row heights
        //    from what the content needed ──
        let mut heights = vec![0.0f32; rows.len()];
        for row in 0..rows.len() {
            for column in 0..columns {
                let Some(cell) = grid[row][column].filter(|c| c.origin == (row, column)) else {
                    continue;
                };
                let width = Self::span_extent(&widths, column, cell.colspan, spacing);
                let height = self.measure_cell(cell, width);
                // A cell that spans rows contributes to none of them directly —
                // its height is shared out below, once the single-row cells
                // have set a floor.
                if cell.rowspan == 1 {
                    heights[row] = heights[row].max(height);
                }
            }
        }
        // A spanning cell only grows the rows it covers if they cannot already
        // hold it, and then evenly — §17.5.3 leaves the distribution to the UA.
        for row in 0..rows.len() {
            for column in 0..columns {
                let Some(cell) = grid[row][column].filter(|c| c.origin == (row, column)) else {
                    continue;
                };
                if cell.rowspan <= 1 {
                    continue;
                }
                let width = Self::span_extent(&widths, column, cell.colspan, spacing);
                let needed = self.measure_cell(cell, width);
                let last = (row + cell.rowspan as usize).min(rows.len());
                let available: f32 = heights[row..last].iter().sum::<f32>()
                    + spacing * ((last - row) as f32 - 1.0);
                if needed > available {
                    let extra = (needed - available) / (last - row) as f32;
                    for height in &mut heights[row..last] {
                        *height += extra;
                    }
                }
            }
        }

        // ── Place ──
        let mut column_x = Vec::with_capacity(columns);
        let mut x = spacing;
        for width in &widths {
            column_x.push(x);
            x += width + spacing;
        }
        let mut row_y = Vec::with_capacity(rows.len());
        let mut y = spacing;
        for height in &heights {
            row_y.push(y);
            y += height + spacing;
        }

        for (row, row_index) in rows.iter().enumerate() {
            for column in 0..columns {
                let Some(cell) = grid[row][column].filter(|c| c.origin == (row, column)) else {
                    continue;
                };
                let width = Self::span_extent(&widths, column, cell.colspan, spacing);
                let last = (row + cell.rowspan as usize).min(rows.len());
                let height: f32 = heights[row..last].iter().sum::<f32>()
                    + spacing * ((last - row) as f32 - 1.0);
                self.place_cell(
                    cell,
                    LayoutRect::new(
                        content_x + column_x[column],
                        content_top + row_y[row],
                        width,
                        height,
                    ),
                );
            }
            // The row box spans the whole table and its own height. It holds
            // the cells in the DOM but positions none of them — a cell that
            // spans rows is in exactly one row's child list and in several
            // rows' worth of space.
            let height = heights[row];
            let rect = LayoutRect::new(
                content_x + spacing,
                content_top + row_y[row],
                (content_w - spacing * 2.0).max(0.0),
                height,
            );
            if let Some(row_box) = self.row_box_mut(*row_index) {
                row_box.set_rect(rect);
            }
        }

        // ── The row GROUPS ────────────────────────────────────────────────
        // A group establishes no formatting context, but it is still a box in
        // the tree: it draws a background and a border, and it has to cover
        // the rows it holds or it draws them nowhere. Its extent is exactly
        // its own rows — which is why it is derived here, after they are
        // placed, rather than being laid out itself.
        //
        // Sizing the group does NOT disturb the rows inside it, even though
        // `set_rect` runs the group's own `relayout()`: a group is given
        // `Formatting::TableRow`, whose relayout does nothing at all. That is
        // not an optimisation — it is what "establishes no formatting context"
        // means, and it is why the table may position boxes that are not its
        // own children without them being re-arranged behind its back.
        let mut extents: Vec<(usize, f32, f32)> = Vec::new();
        for (row, path) in rows.iter().enumerate() {
            let Some(group) = path.group else { continue };
            let (top, bottom) = (row_y[row], row_y[row] + heights[row]);
            match extents.iter_mut().find(|(g, _, _)| *g == group) {
                Some((_, first, last)) => {
                    *first = first.min(top);
                    *last = last.max(bottom);
                }
                None => extents.push((group, top, bottom)),
            }
        }
        for (group, top, bottom) in extents {
            let rect = LayoutRect::new(
                content_x + spacing,
                content_top + top,
                (content_w - spacing * 2.0).max(0.0),
                bottom - top,
            );
            if let Some(box_of_group) = self.children.get_mut(group) {
                box_of_group.set_rect(rect);
            }
        }

        // The table is as tall as its rows, plus the spacing between and around
        // them. Its WIDTH is left alone: `layout_table` is handed a rect and
        // the containing block decides that.
        let used: f32 = heights.iter().sum::<f32>() + spacing * (heights.len() as f32 + 1.0);
        self.content_height = (content_top - r.y) + used + self.padding.bottom;
    }

    /// The table's rows, in VISUAL order, as indices into `self.children`.
    ///
    /// `<tfoot>` renders after `<tbody>` however the markup was written, which
    /// is the whole reason the groups exist. Rows written directly under the
    /// table — legal, and what most hand-written HTML does — count as body
    /// rows.
    ///
    /// Row GROUPS are not returned: they establish nothing, and a table that
    /// treated them as rows would give a whole `<tbody>` one row's height.
    fn table_rows(&mut self) -> Vec<RowPath> {
        // The classification is taken FIRST, as owned values: reaching into a
        // group needs `&mut self` (a `dyn PanelWidget` downcasts only through
        // `as_any_mut`), which cannot be done while iterating `self.children`.
        let kinds: Vec<(usize, String)> = self
            .children
            .iter()
            .enumerate()
            .filter(|(_, child)| !self.out_of_flow.contains(child.name()))
            .map(|(index, child)| {
                let display = self
                    .child_display
                    .get(child.name())
                    .cloned()
                    .unwrap_or_else(|| "table-row".to_string());
                (index, display)
            })
            .collect();

        // §17.2.1 — cells sitting directly in the table share ONE anonymous
        // row, generated before any real row because that is where they are in
        // tree order. Consecutive runs would each get their own row in the full
        // rule; one row is the case that occurs and the one a caller appending
        // cells to a grid means.
        let (mut header, mut body, mut footer) = (Vec::new(), Vec::new(), Vec::new());
        if kinds.iter().any(|(_, display)| display == "table-cell") {
            header.push(RowPath {
                group: None,
                row: ANONYMOUS_ROW,
            });
        }
        for (index, display) in kinds {
            match display.as_str() {
                "table-row" => body.push(RowPath {
                    group: None,
                    row: index,
                }),
                // A group's ROWS are the table's rows. The group box itself is
                // skipped: it is a parent in the tree and nothing in layout —
                // so the walk goes THROUGH it and collects what it holds.
                "table-header-group" => header.extend(self.rows_in_group(index)),
                "table-footer-group" => footer.extend(self.rows_in_group(index)),
                "table-row-group" => body.extend(self.rows_in_group(index)),
                // A caption, a column or a column group. None of them is a row.
                _ => {}
            }
        }
        header.extend(body);
        header.extend(footer);
        header
    }

    /// The rows a `<thead>`/`<tbody>`/`<tfoot>` holds.
    ///
    /// A group's own children are its rows, in tree order — groups do not
    /// reorder among themselves, only against each other. Anything in there
    /// that is not a row is skipped for the same reason it is at table level.
    fn rows_in_group(&mut self, group: usize) -> Vec<RowPath> {
        let Some(box_of_group) = self.children.get_mut(group) else {
            return Vec::new();
        };
        let Some(panel) = box_of_group
            .as_any_mut()
            .and_then(|any| any.downcast_mut::<FlowLayoutPanel>())
        else {
            return Vec::new();
        };
        panel
            .children
            .iter()
            .enumerate()
            .filter(|(_, child)| !panel.out_of_flow.contains(child.name()))
            .filter(|(_, child)| {
                // A group holds rows; anything else in it is not one. The
                // default matches the table-level walk — an unstyled child of a
                // row group is a row, which is what the markup means.
                panel
                    .child_display
                    .get(child.name())
                    .map(String::as_str)
                    .unwrap_or("table-row")
                    == "table-row"
            })
            .map(|(row, _)| RowPath {
                group: Some(group),
                row,
            })
            .collect()
    }

    /// The box holding a row's cells, wherever the row lives.
    ///
    /// The ONE place the group indirection is resolved — every caller works in
    /// `RowPath` and none of them repeats this walk.
    fn row_panel_mut(&mut self, path: RowPath) -> Option<&mut FlowLayoutPanel> {
        // The anonymous row has no box of its own — that is what "anonymous"
        // means. Its cells are the TABLE's own children, so the table stands in
        // for the row it generated, and `cells_of` skips the children that are
        // rows rather than cells.
        if path.row == ANONYMOUS_ROW {
            return Some(self);
        }
        let holder: &mut FlowLayoutPanel = match path.group {
            Some(group) => self
                .children
                .get_mut(group)?
                .as_any_mut()?
                .downcast_mut::<FlowLayoutPanel>()?,
            None => self,
        };
        holder
            .children
            .get_mut(path.row)?
            .as_any_mut()?
            .downcast_mut::<FlowLayoutPanel>()
    }

    /// The row box itself, to be given its rect.
    fn row_box_mut(&mut self, path: RowPath) -> Option<&mut Box<dyn PanelWidget>> {
        // Nothing to give a rect to: an anonymous row is a box the LAYOUT has
        // and the tree does not.
        if path.row == ANONYMOUS_ROW {
            return None;
        }
        match path.group {
            Some(group) => self
                .children
                .get_mut(group)?
                .as_any_mut()?
                .downcast_mut::<FlowLayoutPanel>()?
                .children
                .get_mut(path.row),
            None => self.children.get_mut(path.row),
        }
    }

    /// Build the occupancy grid — which cell sits in which slot, spans included.
    ///
    /// **A cell's column is not its index among its siblings.** A `rowspan` from
    /// an earlier row already occupies a slot, so a later row's second `<td>`
    /// may be the third column. Reconstructing that is the whole job here, and
    /// skipping it is why a naive table renders spans on top of their
    /// neighbours.
    fn build_cell_grid(&mut self, rows: &[RowPath]) -> Vec<Vec<Option<TableCell>>> {
        // Widest row wins, spans counted — a table is as wide as its widest row.
        let mut columns = 0usize;
        for row in rows {
            let width: u32 = self.cells_of(*row).iter().map(|c| c.colspan).sum();
            columns = columns.max(width as usize);
        }
        let mut grid: Vec<Vec<Option<TableCell>>> = vec![vec![None; columns]; rows.len()];
        for (row, row_index) in rows.iter().enumerate() {
            let mut column = 0usize;
            for mut cell in self.cells_of(*row_index) {
                // Step over whatever a row above already claimed.
                while column < columns && grid[row][column].is_some() {
                    column += 1;
                }
                if column >= columns {
                    break;
                }
                let colspan = (cell.colspan as usize).min(columns - column);
                let rowspan = (cell.rowspan as usize).min(rows.len() - row);
                cell.origin = (row, column);
                cell.colspan = colspan as u32;
                cell.rowspan = rowspan as u32;
                for r in row..row + rowspan {
                    for c in column..column + colspan {
                        grid[r][c] = Some(cell);
                    }
                }
                column += colspan;
            }
        }
        grid
    }

    /// The cells of one row, as paths from this table.
    ///
    /// Takes `&mut self` only because the widget trait's child accessor is
    /// `children_mut` — reading a row's children is not a mutation, but there
    /// is no immutable way to ask, and adding one to a shared trait to spare
    /// this signature would be the wrong trade.
    fn cells_of(&mut self, row: RowPath) -> Vec<TableCell> {
        let Some(row_box) = self.row_panel_mut(row) else {
            return Vec::new();
        };
        // **A row's cells are the children that are CELLS.** For a real `<tr>`
        // that is all of them, and an unrecorded display means a cell, which is
        // what the markup means. It matters for the ANONYMOUS row, whose
        // "children" are the table's own: the `<tr>`s among them are rows in
        // their own right and must not be swallowed as cells of the row the
        // table generated around its loose ones.
        //
        // Filtered AFTER `enumerate`, so every cell keeps its real index among
        // its siblings — `with_cell` addresses the child list directly, and a
        // compacted index would reach the wrong box.
        let cells: Vec<(usize, bool)> = row_box
            .children
            .iter()
            .enumerate()
            .map(|(index, child)| {
                let is_cell = !matches!(
                    row_box.child_display.get(child.name()).map(String::as_str),
                    Some("table-row")
                        | Some("table-row-group")
                        | Some("table-header-group")
                        | Some("table-footer-group")
                        | Some("table-column")
                        | Some("table-column-group")
                        | Some("table-caption")
                );
                (index, is_cell)
            })
            .collect();
        let keep: Vec<usize> = cells
            .into_iter()
            .filter(|(_, is_cell)| *is_cell)
            .map(|(index, _)| index)
            .collect();
        row_box
            .children
            .iter_mut()
            .enumerate()
            .filter(|(index, _)| keep.contains(index))
            .map(|(index, child)| {
                // A span lives on the cell, put there by `set_attribute`. A cell
                // that is not a panel — or one nobody gave a span — occupies one
                // slot, which is HTML's own default.
                let (colspan, rowspan) = child
                    .as_any_mut()
                    .and_then(|any| any.downcast_mut::<FlowLayoutPanel>())
                    .map(|panel| (panel.colspan, panel.rowspan))
                    .unwrap_or((1, 1));
                TableCell {
                    row,
                    index,
                    colspan,
                    rowspan,
                    origin: (0, 0),
                }
            })
            .collect()
    }

    /// The width a cell gets: its own column plus every column it spans, and
    /// the spacing that would have run between them.
    fn span_extent(widths: &[f32], column: usize, colspan: u32, spacing: f32) -> f32 {
        let last = (column + colspan as usize).min(widths.len());
        widths[column..last].iter().sum::<f32>() + spacing * ((last - column) as f32 - 1.0)
    }

    /// Reach one cell — a grandchild, through the row that holds it.
    ///
    /// The table positions cells but does not own them: they belong to their
    /// rows, because that is what the markup says and the tree mirrors the
    /// markup. One borrow at a time, which is also why this is a closure rather
    /// than a returned reference.
    fn with_cell<R>(
        &mut self,
        cell: TableCell,
        act: impl FnOnce(&mut Box<dyn PanelWidget>) -> R,
    ) -> Option<R> {
        let row = self.row_panel_mut(cell.row)?;
        Some(act(row.children.get_mut(cell.index)?))
    }

    /// Lay a cell out at a given width and answer the height it needs.
    ///
    /// The cell runs ORDINARY block layout — it is a box like any other, and
    /// giving it its own layout would mean a `<td>` and a `<div>` disagreed
    /// about how to stack their children. Only the WIDTH comes from the table.
    fn measure_cell(&mut self, cell: TableCell, width: f32) -> f32 {
        self.with_cell(cell, |child| {
            let previous = child.rect();
            // A relayout is refused at zero height, so the probe height has to
            // be non-zero — it is replaced by the answer below either way.
            child.set_rect(LayoutRect::new(
                previous.x,
                previous.y,
                width,
                previous.h.max(1.0),
            ));
            match child
                .as_any_mut()
                .and_then(|any| any.downcast_mut::<FlowLayoutPanel>())
            {
                // What the content came out as, which is the whole question a
                // row height asks.
                Some(panel) => panel.content_height,
                // A leaf — a label, an image — is as tall as it is.
                None => child.rect().h,
            }
        })
        .unwrap_or(0.0)
    }

    fn place_cell(&mut self, cell: TableCell, rect: LayoutRect) {
        self.with_cell(cell, |child| child.set_rect(rect));
    }

    /// The width a cell DECLARED, if it declared one.
    ///
    /// The distinction `table-layout: fixed` is built on: §17.5.2.1 sizes a
    /// column from the first row's specified widths and divides what is left
    /// evenly, so a cell that never asked for a width must not be read as
    /// having asked for the one it currently has.
    ///
    /// The fact lives on the ROW, not here — `SetChildWidthMode` announces
    /// each child's `axis_mode` to its own container — so the table asks
    /// through the row that holds the cell, which is the same indirection
    /// every other cell access takes.
    fn cell_declared_width(&mut self, cell: TableCell) -> Option<f32> {
        let name = self.with_cell(cell, |child| child.name().to_string())?;
        let row = self.row_panel_mut(cell.row)?;
        if !row.child_declared_width.contains(&name) {
            return None;
        }
        let width = row.children.get(cell.index)?.rect().w;
        (width > 0.0).then_some(width)
    }

    /// A cell's two intrinsic widths — CSS 2.1 §17.5.2.2 / css-sizing §4.1.
    ///
    /// **min-content** is the widest single word: below it the text has nowhere
    /// left to break. **max-content** is the whole thing on one line. A column
    /// sized between the two is a column that neither clips its content nor
    /// takes more room than it can use, which is the entire argument for a
    /// table over a grid of guessed widths.
    fn cell_intrinsic_widths(&mut self, cell: TableCell) -> (f32, f32) {
        self.with_cell(cell, |child| {
            let Some(panel) = child
                .as_any_mut()
                .and_then(|any| any.downcast_mut::<FlowLayoutPanel>())
            else {
                let w = child.rect().w;
                return (w, w);
            };
            let mut runs: Vec<crate::layout::InlineRun> = Vec::new();
            if !panel.caption.is_empty() {
                runs.push(crate::layout::InlineRun {
                    text: panel.caption.clone(),
                    font: panel.font.clone(),
                    color: panel.colors.foreground,
                    source: None,
                    atomic: None,
                    cursor: None,
                });
            }
            runs.extend(panel.inline_content.iter().filter(|r| r.atomic.is_none()).cloned());

            let (mut min, mut max) = (0.0f32, 0.0f32);
            for run in &runs {
                let (whole, _) = crate::ide_text::measure_rich_text(
                    &[(
                        run.text.clone(),
                        run.font.clone(),
                        cosmic_text::Color::rgb(0, 0, 0),
                    )],
                    None,
                );
                max += whole;
                for word in split_words(&run.text) {
                    let (w, _) = crate::ide_text::measure_rich_text(
                        &[(
                            word.trim_end().to_string(),
                            run.font.clone(),
                            cosmic_text::Color::rgb(0, 0, 0),
                        )],
                        None,
                    );
                    min = min.max(w);
                }
            }
            // A nested box cannot be narrower than it is: it has already been
            // sized, and a column that ignored it would clip it.
            for nested in panel.children.iter() {
                let w = nested.rect().w;
                min = min.max(w);
                max = max.max(w);
            }
            let edges = panel.padding.horizontal();
            (min + edges, max.max(min) + edges)
        })
        .unwrap_or((0.0, 0.0))
    }

    /// Size the columns — §17.5.2, both algorithms.
    ///
    /// `fixed` reads the first row and stops, which is what makes it the fast
    /// one and why a `table-layout: fixed` table can start painting before its
    /// content is known. `auto` measures every cell.
    ///
    /// Cells that SPAN columns take no part in sizing here. §17.5.2.2 leaves
    /// their contribution up to the UA, and the simple readings all distribute
    /// a span's width over columns that single-cell content has already sized
    /// better. Stated rather than silently approximated: a table whose only
    /// content in some column is a spanning cell will size that column from the
    /// other rows, or evenly if there are none.
    fn column_widths(
        &mut self,
        grid: &[Vec<Option<TableCell>>],
        columns: usize,
        cell_area: f32,
    ) -> Vec<f32> {
        let mut widths = vec![0.0f32; columns];
        let even = cell_area / columns as f32;

        if self.table_layout_fixed {
            // The first row alone decides. A cell with no width of its own
            // takes an equal share, which is what a fixed table does with the
            // space nobody claimed.
            //
            // ⚠ A cell's DECLARED width, never its current one. This used to
            // read `child.rect().w` — the width the cell already had — which is
            // a fixed table's own output fed back as its input: laid out once
            // by any other route, `table-layout: fixed` simply kept whatever
            // widths were already there and changed nothing at all. The two
            // algorithms then agreed on every table, which is the one thing
            // §17.5.2.1 exists to prevent.
            for column in 0..columns {
                let width = grid
                    .first()
                    .and_then(|row| row[column])
                    .filter(|c| c.origin == (0, column) && c.colspan == 1)
                    .and_then(|cell| self.cell_declared_width(cell))
                    .unwrap_or(even);
                widths[column] = width;
            }
        } else {
            let mut maxima = vec![0.0f32; columns];
            for row in grid {
                for (column, slot) in row.iter().enumerate().take(columns) {
                    let Some(cell) = slot.filter(|c| c.colspan == 1) else {
                        continue;
                    };
                    if cell.origin.1 != column {
                        continue;
                    }
                    let (min, max) = self.cell_intrinsic_widths(cell);
                    widths[column] = widths[column].max(min);
                    maxima[column] = maxima[column].max(max);
                }
            }
            // Every column starts at min-content, then the slack is shared out
            // towards max-content — §17.5.2.2's "distribute the difference".
            // A column already at its maximum takes none of it.
            let used: f32 = widths.iter().sum();
            let slack = cell_area - used;
            if slack > 0.0 {
                let wanted: f32 = maxima
                    .iter()
                    .zip(&widths)
                    .map(|(max, min)| (max - min).max(0.0))
                    .sum();
                if wanted > 0.0 {
                    let share = (slack / wanted).min(1.0);
                    for column in 0..columns {
                        widths[column] += (maxima[column] - widths[column]).max(0.0) * share;
                    }
                    // Anything still unclaimed — every column at max-content
                    // and room to spare — spreads evenly, which is what makes a
                    // short table fill the width a browser gives it.
                    let left = cell_area - widths.iter().sum::<f32>();
                    if left > 0.0 {
                        for width in &mut widths {
                            *width += left / columns as f32;
                        }
                    }
                } else {
                    for width in &mut widths {
                        *width += slack / columns as f32;
                    }
                }
            } else if used > 0.0 {
                // Narrower than its own minimum: scale rather than overflow,
                // because a column below its min-content width clips text and a
                // proportional squeeze at least keeps the shape.
                let scale = cell_area / used;
                for width in &mut widths {
                    *width = (*width * scale).max(1.0);
                }
            }
        }
        widths
    }

    /// Lay out a box's own text when its children are arranged by some other
    /// formatting context.
    ///
    /// Flex and grid place every child themselves, so nothing there is inline
    /// LEVEL — but the box can still have characters of its own, and they wrap
    /// against the content width exactly as they do in normal flow. Without
    /// this, `inline_items` would be written only by `layout_normal_flow` and
    /// every flex box's caption would vanish the moment paint started reading
    /// from it.
    fn layout_inline_text(&mut self) {
        if self.bordered || (self.caption.is_empty() && self.inline_content.is_empty()) {
            self.inline_items.clear();
            return;
        }
        let content_w = (self.rect.w - self.padding.horizontal()).max(0.0);
        let mut line = LineWriter::new(content_w, self.nowrap, self.text_align);
        let caption = self.caption.clone();
        if !caption.is_empty() {
            line.text_run(&crate::layout::InlineRun {
                text: caption,
                font: self.font.clone(),
                color: self.colors.foreground,
                source: None,
                atomic: None,
                cursor: None,
            });
        }
        for run in self.inline_content.clone() {
            // An atomic slot names a child this formatting context has already
            // positioned. Reserving room for it here would move the text away
            // from a box that is not where the line thinks it is.
            if run.atomic.is_none() {
                line.text_run(&run);
            }
        }
        line.finish();
        self.inline_items = std::mem::take(&mut line.items);
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
        let line_heights: Vec<f32> = if single && !self.cross_size_indefinite() {
            vec![inner_h]
        } else if single {
            // **A row with `height: auto` is as tall as its tallest item** —
            // §9.4. Taking `inner_h` here is what made `align-items: stretch`
            // a fixed point: the items were stretched to the row, the row then
            // measured itself from the items, and it reproduced whatever height
            // it started with — `default_size`'s 150px guess, for a 28px
            // button. The live rect, not the `natural` map: the sizing pass has
            // just settled each item against its own content, and that is the
            // hypothetical cross size this rule asks for.
            vec![
                items
                    .iter()
                    .map(|&i| {
                        let name = self.children[i].name().to_string();
                        let margin = margins.get(&name).copied().unwrap_or_default();
                        self.children[i].rect().h.max(1.0) + margin.vertical()
                    })
                    .fold(0.0f32, f32::max),
            ]
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
        self.content_height = self
            .children
            .iter()
            .map(|c| c.rect().y + c.rect().h)
            .fold(r.y, f32::max)
            - r.y
            + self.padding.bottom;
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
            // Measured ONCE, here, and read back below. The placement loop used
            // to re-derive the base from `child_basis`/`flex`, which was a
            // second copy of `base_main_size` free to disagree with the one the
            // free space was computed from.
            let mut bases: Vec<f32> = Vec::with_capacity(n);
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
                bases.push(base);
                fixed += base;
                let shrink = child_shrink.get(&name).copied().unwrap_or(0.0);
                shrink_scaled += shrink.max(0.0) * base;
            }
            let free = inner_h - gaps - fixed - margin_total;
            // **Nothing is distributed against an indefinite main size** —
            // §9.7. `inner_h` is a guess at this point, so growing into it
            // would hand out space that does not exist and shrinking would
            // squeeze items to fit a height that is about to change. The items
            // stay at their content bases and the container sizes to them.
            let indefinite = self.main_size_indefinite();
            let leftover = if indefinite { 0.0 } else { free.max(0.0) };
            let deficit = if !indefinite && free < 0.0 && shrink_scaled > 0.0 {
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
            for (pos, &i) in line.iter().enumerate() {
                let name = self.children[i].name().to_string();
                let margin = margins.get(&name).copied().unwrap_or_default();
                let base = bases[pos];
                let child = &mut self.children[i];
                let f = child_flex
                    .get(&name)
                    .copied()
                    .unwrap_or_else(|| child.layout_flex());
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
        // **What the flow actually used**, so a container with `height: auto`
        // can size to it. Neither flex layout recorded this, so a flex box kept
        // `default_size`'s 150px guess for ever — and every Flutter `Column`
        // is one.
        self.content_height = self
            .children
            .iter()
            .map(|c| c.rect().y + c.rect().h)
            .fold(r.y, f32::max)
            - r.y
            + self.padding.bottom;
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

    /// Whether the main size this flow distributes is **indefinite** — no
    /// declared extent, so there is nothing yet to divide.
    ///
    /// Block axis only. A block-level box's inline size is always definite:
    /// `width: auto` fills the containing block, and a flex item stretched by
    /// its container has a definite cross size. So a row still divides its
    /// width and only a column defers to content.
    fn main_size_indefinite(&self) -> bool {
        matches!(self.flow_direction, FlowDirection::TopDown) && !self.declared_height
    }

    /// Whether the **cross** size this flow stretches into is indefinite.
    ///
    /// The other axis of the same question: laying out left-to-right the cross
    /// axis is the block one, so an undeclared height means the line's height
    /// comes from its items rather than the items' from the line. A column
    /// crosses the inline axis, which is always definite.
    fn cross_size_indefinite(&self) -> bool {
        matches!(self.flow_direction, FlowDirection::LeftToRight) && !self.declared_height
    }

    /// An item's base size along the main axis, before growing or shrinking.
    ///
    /// A `flex-grow` item contributes **zero**: `flex: 1` means `flex-basis: 0`,
    /// so the item's whole size comes from the free space it is given. Only
    /// non-growing items carry a base into line-breaking.
    fn base_main_size(&self, i: usize) -> f32 {
        let child = self.children[i].as_ref();
        // **An indefinite main size resolves a grower's basis to CONTENT** —
        // §9.2.3. `flex: 1` is `flex: 1 1 0%`, and a percentage of a size that
        // is not known yet is not zero, it is the item's own content. So a
        // column with `height: auto` measures its items and then sizes to their
        // sum; zero here would divide a guess and then re-measure the division.
        //
        // **Every item, growing or not.** `flex-basis: auto` on a non-growing
        // item also means content, and `child_fixed`'s 44px guess is not it: an
        // `AppBar` sibling is `flex: 0`, so a `Scaffold` handed its body 44px
        // and the whole calculator collapsed four levels down from a bar that
        // was not even in the same subtree.
        if self.main_size_indefinite() {
            // Its content, as the last pass left it. **Its live rect, never a
            // maximum with the `natural` map** — that map holds the size the
            // child had when it ARRIVED, which for a container is
            // `default_size`'s 150px guess, and taking the larger of the two
            // pins the item to that guess for ever. The whole point of this
            // branch is that the item has since been measured against its own
            // content, so the remembered size is precisely the number not to
            // trust. `natural` is a fallback for a child not laid out yet, and
            // a child that answers neither keeps the fixed guess rather than
            // collapsing out of view.
            let live = child.rect().h;
            if live > 0.0 {
                return live;
            }
            let natural = self
                .natural
                .get(child.name())
                .map(|(_, h)| *h)
                .unwrap_or(0.0);
            if natural > 0.0 {
                return natural;
            }
            return Self::child_fixed(child);
        }
        // A declared `flex-basis` IS the base, growing or not — that is the
        // difference between "share the space" and "share what is left after
        // my content".
        if let Some(basis) = self.child_basis.get(child.name()) {
            return basis.max(0.0);
        }
        // **A DECLARED main size is the author's answer**, and `child_fixed`'s
        // guess below is not allowed to replace it. A Flutter `AppBar` is
        // `flex: 0` with `height: 56px` and came out 44 — the guess — so the
        // Material bar was the wrong height in every app.
        //
        // Narrow on purpose: only a child whose size the cascade DECLARED,
        // which the container is already told about. Reading the live rect for
        // every non-growing item is the fuller rule and moves margin
        // arithmetic — see the note above `child_fixed`.
        let declared_main = match self.flow_direction {
            FlowDirection::LeftToRight => self.child_declared_width.contains(child.name()),
            FlowDirection::TopDown => self.child_declared_height.contains(child.name()),
        };
        if declared_main {
            let live = match self.flow_direction {
                FlowDirection::LeftToRight => child.rect().w,
                FlowDirection::TopDown => child.rect().h,
            };
            if live > 0.0 {
                return live;
            }
        }
        // ⚠ A non-growing item's base should be `flex-basis: auto` — its
        // CONTENT — and `child_fixed`'s 44px below is a guess instead. It
        // over-charges: a Flutter `Column` reserves 44px each for a status line
        // and a button that measure 30 and 36, and hands the difference to the
        // `Expanded` rows, which then overflow the bottom of the window by
        // about a row.
        //
        // Using the live rect here is the right rule and changes MARGIN
        // arithmetic — `a_margin_pushes_the_element_and_everything_after_it` and
        // `adjacent_margins_do_not_collapse_and_that_is_a_known_divergence` both
        // move — so it is not a one-line swap. Left as the known cause of the
        // remaining overflow rather than taken blind.
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
    /// **A table reaches its cells through the rows that hold them.**
    ///
    /// The trait's child accessor hands back `dyn PanelWidget`, which cannot
    /// answer a `colspan` or a content height — those are this type's own. So
    /// the table downcasts, and this is what makes that possible. The doc on
    /// the trait method says widgets override it "as needed"; a formatting
    /// context that spans two levels of the tree is exactly that need.
    fn as_any_mut(&mut self) -> Option<&mut dyn std::any::Any> {
        Some(self)
    }

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
        } else if !self.inline_items.is_empty() {
            // **Draw what layout decided.** Every word already has a position —
            // see [`LineWriter`] — so this walks the line boxes and paints
            // them; it computes no geometry of its own and cannot disagree with
            // the pass that placed the widgets between them.
            //
            // Positions are relative to the CONTENT origin, so a box that moved
            // without relayout still paints its text inside itself.
            let content_x = r.x + self.padding.left;
            let content_y = r.y + self.top_inset();
            let mut rects: Vec<(String, LayoutRect, Option<crate::css::Cursor>)> = Vec::new();
            for item in &self.inline_items {
                // An atomic slot is a CHILD. It draws itself below, in paint
                // order; the line only ever reserved the room for it.
                if item.atomic.is_some() {
                    continue;
                }
                let (x, y) = (content_x + item.x, content_y + item.y);
                let (cr, cg, cb, ca) = item.color;
                crate::ide_text::draw_rich_text(
                    ctx.pixmap,
                    ctx.font_system,
                    ctx.swash_cache,
                    &[(
                        item.text.clone(),
                        item.font.clone(),
                        cosmic_text::Color::rgba(cr, cg, cb, ca),
                    )],
                    x,
                    y,
                    // No wrap width: the break decisions are already made and
                    // each item is one word on one line. Passing one here would
                    // let the shaper break a word that layout placed whole.
                    None,
                    ctx.scale,
                );
                // **Where each named run landed**, so a click can find it. A run
                // has no widget and no rect, so this is the only geometry an
                // `<a>` dissolved into a line ever gets — and because an item is
                // one word on one line, a link that wraps contributes a rect per
                // line without anyone having to reconstruct where it broke.
                if let Some(source) = &item.source {
                    rects.push((
                        source.clone(),
                        LayoutRect::new(x, y, item.w, item.h),
                        item.cursor,
                    ));
                }
            }
            self.run_rects = rects;
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
        // **A click on a RUN.** Children first, because a widget sitting over
        // the text is in front of it; only then the box's own inline content.
        //
        // Reported as `LinkClicked` for any run, not just an `<a>`: which
        // elements are interactive is the DOM's question — it holds the
        // listeners — and the mapping at the other end already turns this into
        // an ordinary `click` on the named node. A `<span>` with a handler is
        // as clickable as a link, and the painter must not be the thing that
        // decides otherwise.
        if matches!(
            event.kind,
            crate::MouseEventKind::Release(crate::layout::MouseButton::Left)
        ) {
            if let Some((name, ..)) = self
                .run_rects
                .iter()
                .rev()
                .find(|(_, rect, _)| rect.contains(event.x, event.y))
            {
                self.pending_events
                    .push(WidgetEvent::LinkClicked(name.clone()));
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
        // **A run answers for itself.** Same order as the click: children are in
        // front of the text. Reading the run's CASCADED `cursor` is what makes a
        // link in flow show a hand without this ever knowing what a link is —
        // the UA sheet says `a { cursor: pointer }` and every other run keeps
        // the default.
        if let Some((_, _, Some(cursor))) = self
            .run_rects
            .iter()
            .rev()
            .find(|(_, rect, _)| rect.contains(x, y))
        {
            return cursor.icon();
        }
        winit::window::CursorIcon::Default
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        let mut events = std::mem::take(&mut self.pending_events);
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
            // ── Table ────────────────────────────────────────────────────────
            // A span belongs to the CELL and is read by the table; the border
            // model belongs to the table and is declared on it. Both arrive as
            // ordinary declarations, so a page can restyle a table and the
            // layout follows.
            WidgetCommand::Custom(name, value) if name == "SetColspan" || name == "SetRowspan" => {
                if let CommandValue::Number(n) = value {
                    let span = (*n as u32).max(1);
                    if name == "SetColspan" {
                        self.colspan = span;
                    } else {
                        self.rowspan = span;
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetBorderSpacing" => {
                if let Some(px) = command_number(value) {
                    self.border_spacing = px as f32;
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetBorderCollapse" => {
                if let CommandValue::Text(mode) = value {
                    // Collapsing removes the spacing — §17.6.2. The declared
                    // value is kept as it was, so switching back restores it.
                    self.border_collapse = mode.trim().eq_ignore_ascii_case("collapse");
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetTableLayout" => {
                if let CommandValue::Text(mode) = value {
                    self.table_layout_fixed = mode.trim().eq_ignore_ascii_case("fixed");
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetFormatting" => {
                if let CommandValue::Text(mode) = value {
                    let formatting = match mode.trim() {
                        "flex" => Formatting::Flex,
                        "grid" => Formatting::Grid,
                        "table" => Formatting::Table,
                        "table-row" => Formatting::TableRow,
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
            // The same fact as `SetChildHeightMode`, addressed to the box it is
            // ABOUT rather than to its container — a flex container needs to
            // know whether its OWN main size is definite before it can decide
            // what its items' bases resolve against.
            WidgetCommand::Custom(name, value) if name == "SetHeightMode" => {
                if let CommandValue::Text(mode) = value {
                    let declared = mode.trim() == "declared";
                    if declared != self.declared_height {
                        self.declared_height = declared;
                        self.relayout();
                    }
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
                    self.grid_columns = crate::css::parse_track_template(spec).unwrap_or_default();
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetJustifyItems" => {
                if let CommandValue::Text(mode) = value {
                    self.justify_items = mode.trim().to_string();
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetChildJustifySelf" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, mode)) = spec.rsplit_once('=') {
                        self.child_justify_self
                            .insert(child.to_string(), mode.trim().to_string());
                        self.relayout();
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetGridAutoFlow" => {
                if let CommandValue::Text(mode) = value {
                    let m = mode.to_ascii_lowercase();
                    self.grid_flow_column = m.split_whitespace().any(|w| w == "column");
                    self.grid_flow_dense = m.split_whitespace().any(|w| w == "dense");
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value)
                if name == "SetGridAutoRows" || name == "SetGridAutoColumns" =>
            {
                if let CommandValue::Text(spec) = value {
                    if let Some(size) = crate::css::TrackSize::parse(spec) {
                        if name == "SetGridAutoRows" {
                            self.grid_auto_rows = size;
                        } else {
                            self.grid_auto_columns = size;
                        }
                        self.relayout();
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetGridAreas" => {
                if let CommandValue::Text(spec) = value {
                    self.grid_areas = crate::css::parse_area_template(spec).unwrap_or_default();
                    self.relayout();
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetChildGridName" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, area)) = spec.rsplit_once('=') {
                        self.child_grid_name
                            .insert(child.to_string(), area.trim().to_string());
                        self.relayout();
                    }
                }
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetGridRows" => {
                if let CommandValue::Text(spec) = value {
                    self.grid_rows = crate::css::parse_track_template(spec).unwrap_or_default();
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
            // **A container has a text colour too.** `color` was delivered to
            // every box and consumed only by `label.rs`, so a container dropped
            // it: a Flutter `AppBar` declaring `color: #ffffff` painted its
            // title in the default black on its own blue, and nothing in the
            // chain reported a problem — the declaration arrived, reached the
            // widget, and had no reader.
            //
            // It is the CAPTION this colours. An inline child's run carries its
            // own colour from its own computed style, which is why the runs
            // were already right and only the box's own text was wrong.
            WidgetCommand::Custom(name, value) if name == "SetForeColor" => {
                if let Some(rgba) = crate::layout::command_color(value) {
                    self.colors.foreground = rgba;
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
            // `nowrap` is the wrap width being ABSENT — the shaper's own way of
            // saying "one line". `pre`/`pre-wrap`/`pre-line` carry only their
            // wrapping behaviour here; keeping the source's spaces is a
            // question for where text ENTERS the DOM, not for the paint.
            WidgetCommand::Custom(name, CommandValue::Text(value)) if name == "SetWhiteSpace" => {
                self.nowrap = matches!(value.as_str(), "nowrap" | "pre");
                self.relayout();
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

    fn text_run(text: &str) -> crate::layout::InlineRun {
        crate::layout::InlineRun {
            text: text.to_string(),
            font: crate::ide_text::FontSpec::sans(13.0),
            color: (0, 0, 0, 255),
            source: None,
            atomic: None,
            cursor: None,
        }
    }

    /// A table with two rows SEES two rows.
    ///
    /// Isolated from placement deliberately. When four table tests failed at
    /// once, two of them reported the second row's cell still at its
    /// construction size — which is either "the table never reached it" or
    /// "the table reached it and computed nothing", and those have completely
    /// different fixes. This asks the first question on its own.
    #[test]
    fn a_table_sees_every_row_it_holds() {
        let mut table = FlowLayoutPanel::new();
        table.set_rect(LayoutRect::new(0.0, 0.0, 400.0, 300.0));
        table.handle_command(&WidgetCommand::Custom(
            "SetFormatting".into(),
            CommandValue::Text("table".into()),
        ));
        for name in ["r1", "r2"] {
            table.handle_command(&WidgetCommand::Custom(
                "SetChildDisplay".into(),
                CommandValue::Text(format!("{name}=table-row")),
            ));
            let mut row = FlowLayoutPanel::new();
            row.name = name.to_string();
            row.add(Box::new(Button::new(&format!("{name}c1"))));
            table.add(Box::new(row));
        }

        let rows = table.table_rows();
        assert_eq!(rows.len(), 2, "both `<tr>`s are rows of the table");
        // The grid is built from the rows the table FOUND, not from a hand-
        // written index list — otherwise the test asserts against its own
        // guess at the addressing rather than the one layout uses.
        let grid = table.build_cell_grid(&rows);
        assert_eq!(grid.len(), 2, "the grid has a line per row");
        assert_eq!(grid[0].len(), 1, "and a column per cell");
        assert!(
            grid[1][0].is_some(),
            "the second row's cell has a slot of its own"
        );
    }

    /// **The wrapped line** — text, a box, then more text, in a column too
    /// narrow to hold them on one line.
    ///
    /// Both placements got this wrong and for the same reason: the text before
    /// the box was measured with NO wrap width, so a single advance wider than
    /// the column sent the cursor off the content edge. The box then landed at
    /// the START of the second line, painted over the words it should follow,
    /// and the text after it was shaped into what was left — nothing — one
    /// character per line.
    ///
    /// Everything here is asserted against the box's own geometry rather than
    /// against pixel constants: the point is that the three agree, not that the
    /// shaper produces a particular width on a particular font.
    #[test]
    fn a_box_on_a_wrapped_line_follows_the_text_before_it() {
        let mut panel = FlowLayoutPanel::new();
        panel.set_rect(LayoutRect::new(0.0, 0.0, 300.0, 300.0));
        panel.handle_command(&WidgetCommand::Custom(
            "SetFormatting".into(),
            CommandValue::Text("normal".into()),
        ));
        panel.handle_command(&WidgetCommand::Custom(
            "SetChildDisplay".into(),
            CommandValue::Text("field=inline-block".into()),
        ));
        panel.inline_content = vec![
            text_run("This sentence is deliberately long enough that it must wrap before it reaches "),
            crate::layout::InlineRun {
                atomic: Some("field".into()),
                ..text_run("")
            },
            text_run(" and it carries on for a while afterwards."),
        ];
        panel.add(Box::new(Button::new("field")));

        let words: Vec<&InlineItem> = panel
            .inline_items
            .iter()
            .filter(|i| i.atomic.is_none())
            .collect();
        assert!(
            words.len() > 2,
            "the text is split into words so a line can break between them"
        );
        assert!(
            words.iter().any(|w| w.y > 0.0),
            "a sentence this long in a 300px column has to occupy more than one line"
        );
        for word in &words {
            assert!(
                word.x + word.w <= 300.0 + 1.0,
                "a word ran past the content edge at x={} w={}",
                word.x,
                word.w
            );
        }

        // The box sits after the last word of the text before it, on that
        // word's line — not at the start of the line below. Found by POSITION
        // in the item list rather than by matching text, so the assertion does
        // not quietly depend on how `split_words` merges whitespace.
        // An item's position is relative to the CONTENT origin and a child's
        // rect is absolute, so one of them has to move before they can be
        // compared. The child comes to the items, because the items are what
        // the assertions below are about.
        let rect = panel.children[0].rect();
        let box_rect = LayoutRect::new(
            rect.x - (panel.rect.x + panel.padding.left),
            rect.y - (panel.rect.y + panel.top_inset()),
            rect.w,
            rect.h,
        );
        let slot = panel
            .inline_items
            .iter()
            .position(|i| i.atomic.as_deref() == Some("field"))
            .expect("the box takes a slot on the line");
        let before = panel.inline_items[..slot]
            .iter()
            .rfind(|i| i.atomic.is_none())
            .expect("the run before the box was placed");
        assert_eq!(
            box_rect.y, before.y,
            "the box shares a line with the text it follows"
        );
        assert!(
            box_rect.x >= before.x + before.w - 1.0,
            "the box starts where that text ended: box.x={} text ends at {}",
            box_rect.x,
            before.x + before.w
        );

        // And the tail is placed at all. The walk reads runs only when it
        // reaches a box, so everything after the LAST box was consumed by
        // nobody — asserted on the items after the slot, which is empty
        // without the drain rather than merely wrong.
        let tail = &panel.inline_items[slot + 1..];
        assert!(
            !tail.is_empty(),
            "the text after the last box has to be laid out too"
        );
        assert!(
            tail.iter()
                .all(|w| w.y > box_rect.y || w.x >= box_rect.x + box_rect.w - 1.0),
            "the tail follows the box — it does not start back over the top of it"
        );
    }

    /// **The Flutter guard, at the level Flutter actually reaches.**
    ///
    /// `Row`/`Column` map to this widget by KIND — no
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

    /// **The Flutter `Column` guard.** `layout_top_down` was rewritten to
    /// support wrapping, and it is the exact path every Flutter `Column` takes
    /// (by KIND — no element, no cascade). This pins the
    /// unconfigured behaviour: children stack, share the height by flex weight,
    /// and stretch across the cross axis.
    #[test]
    fn an_unconfigured_column_still_stacks_and_stretches() {
        let mut panel = FlowLayoutPanel::new();
        panel.flow_direction = FlowDirection::TopDown;
        panel.set_rect(LayoutRect::new(0.0, 0.0, 400.0, 300.0));
        panel.add(Box::new(Button::new("a")));
        panel.add(Box::new(Button::new("b")));
        panel.add(Box::new(Button::new("c")));

        let r: Vec<LayoutRect> = (0..3).map(|i| panel.child(i).rect()).collect();
        // Stacked, in order, none on top of another.
        assert!(r[0].y < r[1].y && r[1].y < r[2].y, "not stacked: {r:?}");
        // One column — nothing wrapped, because `nowrap` is the default.
        assert!(
            r.iter().all(|x| (x.x - r[0].x).abs() < 0.5),
            "an unconfigured column wrapped: {r:?}"
        );
        // `align-items: stretch` fills the cross axis.
        assert!(
            r[0].w > 300.0,
            "a flex child stopped stretching across: {}",
            r[0].w
        );
        // And each has real height — the flex weights divided the container.
        assert!(r.iter().all(|x| x.h > 10.0), "a child collapsed: {r:?}");
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
    // A `minmax(a, 1fr)` track reserves its FLOOR before anything is shared
    // out, then takes a share of what is left on top. Keeping the floor here
    // is what makes `minmax(200px, 1fr)` never fall below 200px however many
    // tracks compete — the whole point of the function.
    let mut fr_floor = vec![0.0f32; tracks.len()];
    let mut fr_total = 0.0f32;
    let mut used = gaps;
    for (i, track) in tracks.iter().enumerate() {
        let content = auto_sizes.get(i).copied().unwrap_or(0.0).max(0.0);
        match *track {
            TrackSize::Px(v) => sizes[i] = v.max(0.0),
            TrackSize::Percent(p) => sizes[i] = (extent * p / 100.0).max(0.0),
            TrackSize::Auto => sizes[i] = content,
            TrackSize::Fr(f) => {
                fr_total += f.max(0.0);
                continue;
            }
            TrackSize::MinMax(min, max) => {
                let floor = min.definite(extent).unwrap_or(content).max(0.0);
                match max.fr() {
                    // Flexible above its floor: reserve the floor now, share
                    // the leftover below.
                    Some(f) => {
                        fr_total += f.max(0.0);
                        fr_floor[i] = floor;
                        sizes[i] = floor;
                    }
                    // Bounded on both sides: the content, clamped.
                    None => {
                        let cap = max.definite(extent).unwrap_or(f32::INFINITY);
                        sizes[i] = content.clamp(floor, cap.max(floor));
                    }
                }
            }
        }
        used += sizes[i];
    }
    if fr_total > 0.0 {
        let leftover = (extent - used).max(0.0);
        for (i, track) in tracks.iter().enumerate() {
            let weight = match *track {
                TrackSize::Fr(f) => Some(f.max(0.0)),
                TrackSize::MinMax(_, max) => max.fr().map(|f| f.max(0.0)),
                _ => None,
            };
            if let Some(f) = weight {
                sizes[i] = fr_floor[i] + leftover * f / fr_total;
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
    area_names: &std::collections::HashMap<String, String>,
    template: &[Vec<String>],
    col_count: usize,
    dense: bool,
) -> Vec<GridArea> {
    use crate::css::GridLine;
    let cols = col_count.max(1);

    /// The rectangle a named area covers in the template.
    ///
    /// An area is the BOUNDING BOX of every cell bearing its name — that is
    /// how CSS defines it, and it is why `header header` across two columns is
    /// one area two tracks wide rather than two areas.
    fn named_rect(template: &[Vec<String>], name: &str) -> Option<GridArea> {
        let (mut r0, mut c0) = (usize::MAX, usize::MAX);
        let (mut r1, mut c1) = (0usize, 0usize);
        let mut found = false;
        for (r, row) in template.iter().enumerate() {
            for (c, cell) in row.iter().enumerate() {
                if cell == name {
                    found = true;
                    r0 = r0.min(r);
                    c0 = c0.min(c);
                    r1 = r1.max(r + 1);
                    c1 = c1.max(c + 1);
                }
            }
        }
        found.then(|| GridArea {
            col: c0,
            row: r0,
            col_span: c1 - c0,
            row_span: r1 - r0,
        })
    }

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

    // Pass 0: items claiming a NAMED area. These are the most explicit
    // placement there is — the author drew the grid — so they go down first.
    for (i, name) in names.iter().enumerate() {
        let Some(area) = area_names.get(name) else {
            continue;
        };
        let Some(rect) = named_rect(template, area) else {
            continue;
        };
        occupy(rect.col, rect.row, rect.col_span, rect.row_span, &mut occupied);
        areas[i] = Some(rect);
    }

    // Pass 1: everything that named a column line. Explicit first, so the
    // cursor in pass 2 sees them as taken.
    for (i, name) in names.iter().enumerate() {
        if areas[i].is_some() {
            continue;
        }
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
        if dense {
            cursor_col = 0;
            cursor_row = 0;
        }
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
                // **`dense` restarts the search from the top for every
                // item**, so a later small item back-fills a hole an earlier
                // spanning one left. The default is SPARSE: the cursor never
                // moves backwards, which is why a hole stays a hole.
                if dense {
                    cursor_col = 0;
                    cursor_row = 0;
                } else if pinned.is_none() {
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
