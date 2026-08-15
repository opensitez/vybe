//! Layout primitives and the `PanelWidget` trait.
//!
//! This module provides the foundation for the vybe GUI toolkit:
//! - `LayoutRect` — rectangle with hit-testing and subdivision helpers
//! - `MouseEvent` / `KeyEvent` — unified input events
//! - `RenderContext` — bundle of rendering resources
//! - `PanelWidget` trait — implemented by all toolkit panels and containers

use crate::ide_text::FontSpec;
use cosmic_text::{Color, FontSystem, SwashCache};
use std::sync::atomic::{AtomicU64, Ordering};
use tiny_skia::{ColorU8, Pixmap};
use winit::window::CursorIcon;

// ── WidgetId ───────────────────────────────────────────────────────────

/// A unique, lightweight identifier for a widget instance.
///
/// Created via `WidgetId::next()` which returns a globally unique id.
/// Zero is reserved as the "null" / uninitialized sentinel.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct WidgetId(pub u64);

static NEXT_WIDGET_ID: AtomicU64 = AtomicU64::new(1);

impl WidgetId {
    /// The null / uninitialized id.
    pub const NONE: WidgetId = WidgetId(0);

    /// Allocate the next globally-unique widget id.
    pub fn next() -> Self {
        WidgetId(NEXT_WIDGET_ID.fetch_add(1, Ordering::Relaxed))
    }

    /// Whether this is the null id.
    pub fn is_none(self) -> bool {
        self.0 == 0
    }
}

impl Default for WidgetId {
    fn default() -> Self {
        Self::NONE
    }
}

// ── LayoutRect ─────────────────────────────────────────────────────────

/// A rectangle in logical (unscaled) coordinates used for layout.
#[derive(Clone, Copy, Debug, Default)]
pub struct LayoutRect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

impl LayoutRect {
    pub fn new(x: f32, y: f32, w: f32, h: f32) -> Self {
        Self { x, y, w, h }
    }

    pub fn zero() -> Self {
        Self {
            x: 0.0,
            y: 0.0,
            w: 0.0,
            h: 0.0,
        }
    }

    /// Test whether a point lies inside this rectangle.
    pub fn contains(&self, px: f32, py: f32) -> bool {
        px >= self.x && px < self.x + self.w && py >= self.y && py < self.y + self.h
    }

    /// Right edge x coordinate.
    pub fn right(&self) -> f32 {
        self.x + self.w
    }

    /// Bottom edge y coordinate.
    pub fn bottom(&self) -> f32 {
        self.y + self.h
    }

    /// Take a sub-rect from the left edge, shrinking `self`.
    pub fn take_left(&mut self, width: f32) -> LayoutRect {
        let w = width.min(self.w).max(0.0);
        let r = LayoutRect::new(self.x, self.y, w, self.h);
        self.x += w;
        self.w = (self.w - w).max(0.0);
        r
    }

    /// Take a sub-rect from the right edge, shrinking `self`.
    pub fn take_right(&mut self, width: f32) -> LayoutRect {
        let w = width.min(self.w).max(0.0);
        self.w = (self.w - w).max(0.0);
        LayoutRect::new(self.x + self.w, self.y, w, self.h)
    }

    /// Take a sub-rect from the top edge, shrinking `self`.
    pub fn take_top(&mut self, height: f32) -> LayoutRect {
        let h = height.min(self.h).max(0.0);
        let r = LayoutRect::new(self.x, self.y, self.w, h);
        self.y += h;
        self.h = (self.h - h).max(0.0);
        r
    }

    /// Take a sub-rect from the bottom edge, shrinking `self`.
    pub fn take_bottom(&mut self, height: f32) -> LayoutRect {
        let h = height.min(self.h).max(0.0);
        self.h = (self.h - h).max(0.0);
        LayoutRect::new(self.x, self.y + self.h, self.w, h)
    }
}

// ── Input Events ───────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum MouseButton {
    Left,
    Right,
    Middle,
}

#[derive(Clone, Copy, Debug)]
pub enum MouseEventKind {
    Press(MouseButton),
    Release(MouseButton),
    Move,
    Scroll(f32),
}

/// A mouse event with absolute logical coordinates.
#[derive(Clone, Copy, Debug)]
pub struct MouseEvent {
    pub x: f32,
    pub y: f32,
    pub kind: MouseEventKind,
    pub cmd: bool,
    pub shift: bool,
    pub alt: bool,
}

/// A keyboard event with modifier state.
#[derive(Clone, Debug)]
pub struct KeyEvent {
    pub logical_key: winit::keyboard::Key,
    pub key_without_modifiers: winit::keyboard::Key,
    pub state: winit::event::ElementState,
    pub cmd: bool,
    pub shift: bool,
    pub alt: bool,
    pub text: Option<String>,
}

// ── Render Context ─────────────────────────────────────────────────────

/// Bundle of resources needed for rendering.
pub struct RenderContext<'a> {
    pub pixmap: &'a mut Pixmap,
    pub font_system: &'a mut FontSystem,
    pub swash_cache: &'a mut SwashCache,
    pub scale: f32,
}

impl<'a> RenderContext<'a> {
    /// Draw the toolkit's default UI text at physical pixel coordinates.
    ///
    /// The style is a [`FontSpec`] like any other — this one just happens to
    /// carry the defaults this helper has always used. Callers with a declared
    /// style call [`RenderContext::draw_text_styled`] instead of having their
    /// font decided for them here.
    pub fn draw_text(&mut self, text: &str, x: f32, y: f32, r: u8, g: u8, b: u8, a: u8) {
        let spec = FontSpec::mono(14.0).with_line_height(20.0);
        self.draw_text_styled(text, x, y, &spec, r, g, b, a);
    }

    /// Draw one line in a fully specified style, at physical pixel coordinates.
    pub fn draw_text_styled(
        &mut self,
        text: &str,
        x: f32,
        y: f32,
        spec: &FontSpec,
        r: u8,
        g: u8,
        b: u8,
        a: u8,
    ) {
        crate::ide_text::draw_text_spec_physical(
            self.pixmap,
            self.font_system,
            self.swash_cache,
            text,
            x,
            y,
            spec,
            Color::rgba(r, g, b, a),
            self.scale,
        );
    }

    /// Draw monospace UI text with a custom font size at physical pixel coordinates.
    pub fn draw_text_sized(
        &mut self,
        text: &str,
        x: f32,
        y: f32,
        font_size: f32,
        r: u8,
        g: u8,
        b: u8,
        a: u8,
    ) {
        // 1.5, not the module default of 1.3 — this helper has always used a
        // looser line, and the line box is what a baseline sits inside.
        let spec = FontSpec::mono(font_size).with_line_height(font_size * 1.5);
        self.draw_text_styled(text, x, y, &spec, r, g, b, a);
    }
}

// ── Cursor Motion ──────────────────────────────────────────────────────

/// Direction for cursor movement — abstracts cosmic_text::Motion.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum CursorMotion {
    Left,
    Right,
    Up,
    Down,
    Home,
    End,
    BufferStart,
    BufferEnd,
    LeftWord,
    RightWord,
}

// ── Dock ───────────────────────────────────────────────────────────────

/// Dock position for children of a `DockPanel`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Dock {
    Left,
    Right,
    Top,
    Bottom,
    Fill,
}

// ── Widget Events ──────────────────────────────────────────────────────

/// A dynamic value that can be exchanged via commands.
#[derive(Clone, Debug)]
pub enum CommandValue {
    /// No value / not applicable.
    None,
    /// A string value (text, label, path, etc.).
    Text(String),
    /// A numeric value (slider position, counter value, etc.).
    Number(f64),
    /// A boolean value (checked, enabled, visible, etc.).
    Bool(bool),
    /// An index value (selected item, tab index, etc.).
    Index(usize),
    /// An RGBA colour.
    Color(u8, u8, u8, u8),
    /// A box's inline content — see [`InlineRun`].
    Runs(Vec<InlineRun>),
}

/// **A styled span of a box's text** — wxhtmledit's `InlineRun`.
///
/// The thing that makes `a <strong>b</strong> c` one line of mixed text rather
/// than a paragraph with a block stacked underneath it. An inline element is
/// **not a box**: it contributes a differently-styled run of its parent's text,
/// which is why `<strong>` has no rect, no padding and no position of its own,
/// and why asking a toolkit for "the strong widget" is the wrong question.
///
/// The style is resolved — the cascade has already run — because a run is
/// produced from a computed style rather than from declarations. That is the
/// same reason wxhtmledit propagates `box.style` into `inlineContent` right
/// after `InheritFromParent`.
#[derive(Clone, Debug, PartialEq)]
pub struct InlineRun {
    pub text: String,
    pub font: crate::ide_text::FontSpec,
    pub color: (u8, u8, u8, u8),
}

/// Read a `Custom` command payload as a number. Commands cross the host
/// boundary as text (`set_property` stringifies every value), so a numeric
/// payload arrives as either `Number` or a `Text` that parses.
pub fn command_number(val: &CommandValue) -> Option<f64> {
    match val {
        CommandValue::Number(n) => Some(*n),
        CommandValue::Index(i) => Some(*i as f64),
        CommandValue::Text(s) => s.trim().parse::<f64>().ok(),
        _ => None,
    }
}

/// The order children are PAINTED in — CSS 2.1 Appendix E, the parts of it a
/// toolkit needs.
///
/// **Three buckets, not two.** A positioned box with a NEGATIVE `z-index`
/// paints *below* the in-flow content, not above it with the other positioned
/// boxes — that is the whole point of `z-index: -1`, and getting it wrong makes
/// a "send behind" impossible to express. (wxhtmledit's `CollectPaintOrder`
/// has exactly two buckets and this exact bug.)
///
///   1. positioned, `z-index` < 0   — behind everything
///   2. everything in normal flow   — document order
///   3. positioned, `z-index` >= 0  — sorted, stably
///
/// A **stable** sort throughout, so equal `z-index` keeps document order and
/// the default — no `z-index` declared anywhere — is exactly document order.
///
/// `is_positioned` is asked rather than assumed because "positioned" means
/// `position` is not `static`, which only the caller knows: for a flow
/// container it is the out-of-flow/relative sets, and for a form whose children
/// all carry coordinates it is simply true.
pub fn paint_order(
    count: usize,
    is_positioned: impl Fn(usize) -> bool,
    z_of: impl Fn(usize) -> i32,
) -> Vec<usize> {
    let mut below = Vec::new();
    let mut flow = Vec::new();
    let mut above = Vec::new();
    for i in 0..count {
        if !is_positioned(i) {
            flow.push(i);
        } else if z_of(i) < 0 {
            below.push(i);
        } else {
            above.push(i);
        }
    }
    below.sort_by_key(|i| z_of(*i));
    above.sort_by_key(|i| z_of(*i));
    below.extend(flow);
    below.extend(above);
    below
}

#[cfg(test)]
mod color_tests {
    use super::parse_color;

    #[test]
    fn the_widget_path_and_the_css_path_are_one_parser() {
        // These were two implementations with two grammars. Each line below
        // answered `None` here and a colour in `css.rs` — a value's meaning
        // decided by which of two functions it happened to reach.
        assert_eq!(parse_color("#f00"), Some((255, 0, 0, 255)), "CSS shorthand");
        assert_eq!(
            parse_color("rgb(255, 0, 0)"),
            Some((255, 0, 0, 255)),
            "rgb() is CSS's own syntax and was unparseable here"
        );
        assert_eq!(parse_color("navy"), Some((0, 0, 128, 255)), "a CSS colour");
        assert_eq!(parse_color("rgba(0, 0, 255, 0.5)").map(|c| c.3), Some(128));
    }

    #[test]
    fn eight_digit_hex_is_rgba_which_is_a_deliberate_behaviour_change() {
        // The contradiction, not a gap: `#RRGGBBAA` was read as RGBA by
        // `css.rs` and as ARGB here, so `#FF0000FF` was RED down one path and
        // BLUE down the other. CSS Color 4 says the alpha is LAST, so CSS
        // wins and this path changed.
        assert_eq!(
            parse_color("#FF0000FF"),
            Some((255, 0, 0, 255)),
            "alpha is the LAST byte — this used to answer blue"
        );
        assert_eq!(parse_color("#00FF0080").map(|c| c.3), Some(128));
    }

    #[test]
    fn a_packed_integer_is_still_argb_because_the_toolkits_say_so() {
        // The other channel order survives as its own case rather than as a
        // clash: Flutter's `Color.value`, WinForms' `Color.ToArgb` and VCL's
        // `TColor` all hand over ARGB, and CSS has no syntax that collides.
        assert_eq!(parse_color("0xFF0000FF"), Some((0, 0, 255, 255)), "ARGB");
        // No alpha byte means opaque, not invisible.
        assert_eq!(parse_color("0x2196F3").map(|c| c.3), Some(255));
        assert_eq!(parse_color("4294901760"), Some((255, 0, 0, 255)));
    }

    #[test]
    fn the_toolkit_only_colour_names_survived_the_move() {
        assert_eq!(parse_color("lightgray"), Some((211, 211, 211, 255)));
        assert_eq!(parse_color("darkgrey"), Some((169, 169, 169, 255)));
        assert_eq!(parse_color("transparent"), Some((0, 0, 0, 0)));
        assert_eq!(parse_color("nonsense"), None);
    }
}

#[cfg(test)]
mod paint_order_tests {
    use super::paint_order;

    #[test]
    fn a_negative_z_index_paints_behind_the_normal_flow() {
        // THE rule two buckets get wrong. Index 0 is positioned at z = -1, 1 is
        // in flow, 2 is positioned at z = 0. A "flow first, then positioned"
        // split puts 0 on top of 1; CSS puts it underneath, which is the only
        // thing `z-index: -1` is for.
        let positioned = |i: usize| i != 1;
        let z = |i: usize| if i == 0 { -1 } else { 0 };
        assert_eq!(paint_order(3, positioned, z), vec![0, 1, 2]);
    }

    #[test]
    fn equal_z_keeps_document_order_so_declaring_nothing_changes_nothing() {
        let order = paint_order(4, |_| true, |_| 0);
        assert_eq!(order, vec![0, 1, 2, 3]);
    }

    #[test]
    fn positioned_boxes_paint_over_the_flow_whatever_the_tree_says() {
        // 0 is positioned and FIRST in the tree; 1 and 2 are in flow. The
        // positioned box still paints last — a `position: relative` with no
        // offset covering its siblings is the same rule.
        let order = paint_order(3, |i| i == 0, |_| 0);
        assert_eq!(order, vec![1, 2, 0]);
    }
}

/// Read a `Custom` command payload as four per-side edges.
///
/// A bare number is uniform — that is CSS's own one-value shorthand, and it is
/// what every caller predating per-side padding sends. Four comma-separated
/// numbers are `top,right,bottom,left`, CSS order, so a caller that has already
/// resolved a shorthand does not have to re-serialise it as CSS text and have
/// it re-parsed here.
pub fn command_edges(val: &CommandValue) -> Option<crate::css::Edges> {
    if let CommandValue::Text(s) = val {
        let parts: Vec<&str> = s.split(',').collect();
        if parts.len() == 4 {
            let side = |i: usize| parts[i].trim().parse::<f32>().ok();
            return Some(crate::css::Edges {
                top: side(0)?,
                right: side(1)?,
                bottom: side(2)?,
                left: side(3)?,
            });
        }
    }
    command_number(val).map(|n| crate::css::Edges::uniform(n as f32))
}

/// Read a `Custom` command payload as an RGBA colour: a `Color` payload, a
/// `#RRGGBB`/`#AARRGGBB` string, a named colour, or a packed ARGB integer
/// (what Flutter's `Color(0xFF2196F3)` carries).
pub fn command_color(val: &CommandValue) -> Option<(u8, u8, u8, u8)> {
    match val {
        CommandValue::Color(r, g, b, a) => Some((*r, *g, *b, *a)),
        CommandValue::Number(n) => {
            Some(argb_u32_to_rgba(crate::css::opaque_if_no_alpha(*n as u32)))
        }
        CommandValue::Text(s) => parse_color(s),
        _ => None,
    }
}

/// Split a packed `0xAARRGGBB` integer into RGBA components.
///
/// Alpha comes out verbatim. The "a six-digit constant means opaque, not
/// invisible" rule belongs where the *syntax* is read
/// (`css::opaque_if_no_alpha`), because only there is it known that the input
/// had no alpha byte at all. Applying it here instead made `transparent`
/// inexpressible: it parses to `0x00000000`, and this promoted it back to
/// opaque black on the way out.
fn argb_u32_to_rgba(v: u32) -> (u8, u8, u8, u8) {
    let a = ((v >> 24) & 0xFF) as u8;
    let r = ((v >> 16) & 0xFF) as u8;
    let g = ((v >> 8) & 0xFF) as u8;
    let b = (v & 0xFF) as u8;
    (r, g, b, a)
}

/// Parse a colour string into RGBA channels.
///
/// **One parser, in `css.rs`, where a CSS value belongs.** This was a second
/// implementation with a different grammar, and the two disagreed on real
/// input:
///
/// - `#f00` — CSS shorthand. Parsed there, `None` here.
/// - `rgb()` / `rgba()` — parsed there, `None` here.
/// - Ten named colours (`silver`, `maroon`, `lime`, `navy`, `teal`, `aqua`, …)
///   — known there, `None` here.
/// - **`#RRGGBBAA` — read as RGBA there and as ARGB here.** Not a gap, a
///   contradiction: `#FF0000FF` was red down one path and blue down the other,
///   decided by which of two functions a value happened to reach. CSS is
///   right, so CSS wins, and the toolkits' packed-integer form (which really
///   is ARGB) moved across as its own case rather than as a clash.
pub fn parse_color(s: &str) -> Option<(u8, u8, u8, u8)> {
    crate::css::parse_color(s).map(argb_u32_to_rgba)
}

/// Tri-state check state for checkboxes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CheckState {
    Unchecked,
    Checked,
    Indeterminate,
}

impl CheckState {
    /// Cycle to the next state on toggle: Unchecked -> Checked -> Unchecked.
    /// Indeterminate is only set programmatically and cycles to Checked on toggle.
    pub fn toggle(self) -> Self {
        match self {
            CheckState::Unchecked => CheckState::Checked,
            CheckState::Checked => CheckState::Unchecked,
            CheckState::Indeterminate => CheckState::Checked,
        }
    }

    pub fn is_checked(self) -> bool {
        self == CheckState::Checked
    }
}

/// Selection mode for list-based controls.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SelectionMode {
    /// Only one item can be selected at a time (default).
    Single,
    /// Multiple items can be toggled independently with simple clicks.
    MultiSimple,
    /// Extended multi-select: Ctrl+click toggles, Shift+click selects range.
    MultiExtended,
}

/// Text alignment for labels and similar widgets.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextAlign {
    Left,
    Center,
    Right,
}

impl TextAlign {
    /// A CSS `text-align` keyword, or the `SetTextAlign` payload that carries
    /// one.
    ///
    /// `start`/`end` are the logical spellings and mean left/right in a
    /// left-to-right document, which is the only direction the toolkit lays
    /// out; `justify` has no line breaker behind it and reads as `left` rather
    /// than as a silently different alignment.
    ///
    /// Stated once, here, because a widget that parsed the keyword itself is
    /// how three controls end up disagreeing about what `center` means.
    pub fn from_css(value: &str) -> Option<TextAlign> {
        match value.trim().to_ascii_lowercase().as_str() {
            "left" | "start" | "justify" => Some(TextAlign::Left),
            "center" | "centre" => Some(TextAlign::Center),
            "right" | "end" => Some(TextAlign::Right),
            _ => None,
        }
    }
}

// ── Anchor ──────────────────────────────────────────────────────────────

/// Anchor edges — bitflags indicating which edges of the parent the widget
/// is anchored to. When a parent resizes, anchored edges maintain their
/// distance to the parent edge. Un-anchored edges allow the widget to grow/shrink.
///
/// Default is `TOP | LEFT` (widget stays at its position, doesn't resize).
/// Set `TOP | LEFT | RIGHT` to stretch horizontally with parent.
/// Set all four to stretch in both axes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Anchor(pub u8);

impl Anchor {
    pub const NONE: Anchor = Anchor(0);
    pub const TOP: Anchor = Anchor(1);
    pub const BOTTOM: Anchor = Anchor(2);
    pub const LEFT: Anchor = Anchor(4);
    pub const RIGHT: Anchor = Anchor(8);
    pub const TOP_LEFT: Anchor = Anchor(1 | 4);
    pub const ALL: Anchor = Anchor(1 | 2 | 4 | 8);

    pub fn has_top(self) -> bool {
        self.0 & 1 != 0
    }
    pub fn has_bottom(self) -> bool {
        self.0 & 2 != 0
    }
    pub fn has_left(self) -> bool {
        self.0 & 4 != 0
    }
    pub fn has_right(self) -> bool {
        self.0 & 8 != 0
    }
}

impl std::ops::BitOr for Anchor {
    type Output = Anchor;
    fn bitor(self, rhs: Anchor) -> Anchor {
        Anchor(self.0 | rhs.0)
    }
}

impl Default for Anchor {
    fn default() -> Self {
        Anchor::TOP_LEFT
    }
}

/// Stored state for anchor layout: the widget's position relative to its parent
/// at the time of initial placement, plus the parent size at that time.
#[derive(Clone, Copy, Debug, Default)]
pub struct AnchorLayout {
    /// Widget rect when anchor was first set (relative to parent origin).
    pub initial_rect: LayoutRect,
    /// Parent size when anchor was first set.
    pub parent_size: (f32, f32),
    /// Anchor edges.
    pub anchor: Anchor,
}

impl AnchorLayout {
    /// Compute where the widget should be placed given the new parent size.
    pub fn resolve(
        &self,
        new_parent_w: f32,
        new_parent_h: f32,
        parent_x: f32,
        parent_y: f32,
    ) -> LayoutRect {
        let ir = self.initial_rect;
        let (pw, ph) = self.parent_size;
        if pw <= 0.0 || ph <= 0.0 {
            return LayoutRect::new(parent_x + ir.x, parent_y + ir.y, ir.w, ir.h);
        }
        let a = self.anchor;

        // Distances from edges in original parent
        let dist_left = ir.x;
        let dist_top = ir.y;
        let dist_right = pw - (ir.x + ir.w);
        let dist_bottom = ph - (ir.y + ir.h);

        // Compute new x & w
        let (nx, nw) = if a.has_left() && a.has_right() {
            // Both: maintain distance from both edges → widget stretches
            let l = dist_left;
            let r = new_parent_w - dist_right;
            (l, (r - l).max(0.0))
        } else if a.has_right() {
            // Right only: maintain distance from right edge
            let r = new_parent_w - dist_right;
            (r - ir.w, ir.w)
        } else {
            // Left (default): maintain distance from left edge
            (dist_left, ir.w)
        };

        // Compute new y & h
        let (ny, nh) = if a.has_top() && a.has_bottom() {
            let t = dist_top;
            let b = new_parent_h - dist_bottom;
            (t, (b - t).max(0.0))
        } else if a.has_bottom() {
            let b = new_parent_h - dist_bottom;
            (b - ir.h, ir.h)
        } else {
            (dist_top, ir.h)
        };

        LayoutRect::new(parent_x + nx, parent_y + ny, nw, nh)
    }
}

/// Apply anchor layout to a list of (AnchorLayout, widget) pairs when the parent resizes.
pub fn apply_anchor_layouts(
    layouts: &[AnchorLayout],
    widgets: &mut [Box<dyn PanelWidget>],
    parent_rect: LayoutRect,
) {
    for (al, w) in layouts.iter().zip(widgets.iter_mut()) {
        if al.anchor == Anchor::NONE {
            continue;
        }
        let new_rect = al.resolve(parent_rect.w, parent_rect.h, parent_rect.x, parent_rect.y);
        w.set_rect(new_rect);
    }
}

/// Commands sent from the host application **to** a widget.
#[derive(Clone, Debug)]
pub enum WidgetCommand {
    /// Set a layout flex weight (0 = fixed/natural size, >0 = share leftover
    /// space). Used by the Flutter adapter: a Scaffold app-bar is fixed, the
    /// body flexes; `Expanded`/`Flexible` flex, plain content is fixed.
    SetFlex(f32),
    /// Set the widget's text / label content.
    SetText(String),
    /// Get the widget's current text.
    GetText,
    /// Set a numeric value (slider position, progress, numeric up-down, etc.).
    SetValue(f64),
    /// Get the widget's current value.
    GetValue,
    /// Enable or disable the widget.
    SetEnabled(bool),
    /// Show or hide the widget.
    SetVisible(bool),
    /// Request focus for this widget.
    Focus,
    /// Set the selected index (list, combo, tabs, etc.).
    SetSelectedIndex(usize),
    /// Set checked / selected state (checkbox, radio).
    SetChecked(bool),
    /// Append an item to the widget's list.
    AddItem(String),
    /// Remove item at index.
    RemoveItem(usize),
    /// Read the item at index — `select.options[i].text`, `TStrings[i]`,
    /// .NET's `this[int]`. Answers `CommandValue::Text`, or `None` when the
    /// index is out of range.
    ///
    /// The list already had add, remove and clear but **no read**, so an
    /// indexed item was unreachable from any frontend. A control verb is
    /// `[control]` in and one value out, which is why an index needs its own
    /// command rather than riding on `Custom`.
    GetItem(usize),
    /// Replace the item at index. The write half of the same pair — declared
    /// together because an indexer needs both directions before the
    /// compiler's `declared_indexer_emits` will take the branch at all.
    SetItem(usize, String),
    /// Remove all items.
    ClearItems,
    /// Custom command with a string key and arbitrary payload.
    Custom(String, CommandValue),
}

/// Events emitted by widgets back to the host application.
///
/// Containers collect these; the host matches on them to handle callbacks
/// without knowing widget layout details.
#[derive(Clone, Debug)]
pub enum WidgetEvent {
    /// A tab was selected. Payload: tab index.
    TabChanged(usize),
    /// A tab close button was clicked. Payload: tab index.
    TabCloseRequested(usize),
    /// A status-bar section was clicked. Payload: section click_id.
    StatusBarClick(String),
    /// A tree-view item was opened. Payload: file path.
    TreeItemOpened(String),
    /// A dropdown item was selected. Payload: dropdown id, item index.
    DropdownSelected(String, usize),
    /// A menu action was selected. Payload: action id string.
    MenuAction(String),
    /// A button was clicked. Payload: button name.
    ButtonClicked(String),
    /// A checkbox was toggled. Payload: checkbox name, new checked state.
    CheckboxToggled(String, bool),
    /// A text field value changed. Payload: field name, new text.
    TextChanged(String, String),
    /// A color was picked. Payload: widget name, hex string (e.g. "#FF0000").
    ColorChanged(String, String),
    /// A radio button was selected. Payload: name, selected state.
    RadioSelected(String, bool),
    /// A slider value changed. Payload: name, new value (0.0..1.0).
    SliderChanged(String, f32),
    /// A list box item was selected. Payload: name, selected index.
    ListBoxSelected(String, usize),
    /// A list view item was selected. Payload: name, selected index.
    ListViewSelected(String, usize),
    /// A numeric up-down value changed. Payload: name, new value.
    NumericChanged(String, f64),
    /// A scroll bar position changed. Payload: name, position (0.0..1.0).
    ScrollChanged(String, f32),
    /// A calendar date was selected. Payload: name, day of month.
    CalendarDateSelected(String, u32),
    /// A menu strip item was clicked. Payload: name, item index.
    MenuItemClicked(String, usize),
    /// A context menu item was clicked. Payload: name, item index.
    ContextMenuItemClicked(String, usize),
    /// A tool strip button was clicked. Payload: name, item index.
    ToolStripItemClicked(String, usize),
    /// A split container divider was moved. Payload: name, new position (0.0..1.0).
    SplitMoved(String, f32),
    /// A link label was clicked. Payload: name.
    LinkClicked(String),
    /// A select/combobox item was chosen. Payload: name, selected index.
    SelectChanged(String, usize),
    /// A tab control tab was selected. Payload: name, tab index.
    TabControlChanged(String, usize),
    /// Generic named action (for custom widgets).
    Action(String),
    /// Mouse entered a widget. Payload: widget name.
    MouseEnter(String),
    /// Mouse left a widget. Payload: widget name.
    MouseLeave(String),
}

// ── PanelWidget Trait ──────────────────────────────────────────────────

/// Core trait for widgets in the toolkit layout system.
///
/// All panel containers and leaf widgets implement this trait.
/// Coordinates are in logical (unscaled) pixels.
pub trait PanelWidget: Send + Sync {
    /// Assign a layout rectangle to this widget.
    fn set_rect(&mut self, rect: LayoutRect);

    /// Get the current layout rectangle.
    fn rect(&self) -> LayoutRect;

    /// Recursively search for a widget by name and return its layout rect.
    fn find_rect(&self, name: &str) -> Option<LayoutRect> {
        if self.name() == name {
            Some(self.rect())
        } else {
            None
        }
    }

    /// This widget's own children, mutably. Default empty — "I'm a leaf."
    /// Containers override; that one override is what lets the free functions
    /// [`find_widget_mut`] and [`take_widget`] walk the whole tree, so a node
    /// stays reachable and mutable by name however deeply it is nested.
    /// That reachability is what makes the tree a document rather than a
    /// display list.
    fn children_mut(&mut self) -> Vec<&mut Box<dyn PanelWidget>> {
        Vec::new()
    }

    /// Remove a DIRECT child by name and hand it back. Default `None` — "not
    /// a container." This is `removeChild`, and the "move" half of
    /// `appendChild`: a node has one parent, so inserting it elsewhere takes
    /// it out of where it was. Use [`take_widget`] to search a whole subtree.
    fn detach(&mut self, _name: &str) -> Option<Box<dyn PanelWidget>> {
        None
    }

    /// Render the widget into the pixmap.
    fn render(&mut self, ctx: &mut RenderContext);

    /// Downcast hook for the host bridge.
    ///
    /// Returns `Some(&mut dyn Any)` so callers can `downcast_mut` to a
    /// specific widget type. Default-`None` so existing widgets don't
    /// need updating; widgets the host bridge needs to look up
    /// (currently just `Canvas`) override this to return
    /// `Some(self as &mut dyn Any)`.
    fn as_any_mut(&mut self) -> Option<&mut dyn std::any::Any> {
        None
    }

    /// Nest a child widget into this container and re-run its layout.
    ///
    /// Default returns `Some(child)` — "I'm not a container, here it is back."
    /// Layout containers (FlowLayoutPanel/StackPanel/Panel) override to take
    /// ownership, arrange, and return `None`. This is how the host bridge lets
    /// vybe_widgets own the widget tree + layout instead of flat-positioning.
    fn add_child(&mut self, child: Box<dyn PanelWidget>) -> Option<Box<dyn PanelWidget>> {
        Some(child)
    }

    /// Nest `child` at `index` among this container's children.
    ///
    /// Default REFUSES by handing the child back, exactly as `add_child` does
    /// — and deliberately does not fall back to appending. `insertBefore`
    /// differs from `appendChild` only in where the child lands, so a
    /// container that quietly appended would leave the DOM reading back one
    /// order while the window showed another, which is the least findable
    /// class of bug this toolkit has.
    fn insert_child(
        &mut self,
        _index: usize,
        child: Box<dyn PanelWidget>,
    ) -> Option<Box<dyn PanelWidget>> {
        Some(child)
    }

    /// Route a name-targeted command down the tree. `None` = "no widget here
    /// owns that name" (keep searching siblings); `Some(_)` = handled. Leaf
    /// widgets match their own name; containers override to recurse into
    /// children so nested controls (a Flutter tree) stay reachable by name —
    /// the basis for `setState` updating a control's property in place.
    fn send_command_named(&mut self, name: &str, cmd: &WidgetCommand) -> Option<CommandValue> {
        if self.name() == name {
            Some(self.handle_command(cmd))
        } else {
            None
        }
    }

    /// Layout flex weight when this widget is a child of a flex container
    /// (`FlowLayoutPanel`). 0 = fixed/natural size; >0 = share leftover space.
    /// Default 1 (flex). Containers that carry a set weight override this.
    fn layout_flex(&self) -> f32 {
        1.0
    }

    /// Add `child` into the container named `parent_name`, searching this
    /// widget and its descendants. Returns `None` if it was placed, or
    /// `Some(child)` (unconsumed) if `parent_name` wasn't found here — the
    /// top-down counterpart to `add_child`, used by the Flutter realizer which
    /// creates parents before children.
    fn add_child_to(
        &mut self,
        parent_name: &str,
        child: Box<dyn PanelWidget>,
    ) -> Option<Box<dyn PanelWidget>> {
        if self.name() == parent_name {
            self.add_child(child)
        } else {
            Some(child)
        }
    }

    /// Handle a mouse event. Returns `true` if the event was consumed.
    fn handle_mouse(&mut self, event: &MouseEvent) -> bool;

    /// Handle a key event. Returns `true` if the event was consumed.
    fn handle_key(&mut self, event: &KeyEvent) -> bool;

    /// Handle a scroll event at (x, y) with the given delta.
    /// Returns `true` if the event was consumed.
    fn handle_scroll(&mut self, _delta: f32, _x: f32, _y: f32) -> bool {
        false
    }

    /// Return the desired cursor icon for the given position.
    fn cursor_at(&self, _x: f32, _y: f32) -> CursorIcon {
        CursorIcon::Default
    }

    /// Drain any pending events from this widget (and its children).
    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        Vec::new()
    }

    /// Whether this widget wants keyboard focus.
    fn focusable(&self) -> bool {
        false
    }

    /// Widget name (used by FocusManager for identification).
    fn name(&self) -> &str {
        ""
    }

    /// Notify the widget that it gained or lost focus.
    fn set_focused(&mut self, _focused: bool) {}

    /// Whether the mouse is currently hovering over this widget.
    fn hovered(&self) -> bool {
        false
    }

    /// Notify the widget that the mouse entered or left it.
    fn set_hovered(&mut self, _hovered: bool) {}

    /// Unique identifier for this widget instance.
    fn widget_id(&self) -> WidgetId {
        WidgetId::NONE
    }

    /// Handle a command sent from the host application.
    /// Returns a `CommandValue` for queries (`GetText`, `GetValue`), or `CommandValue::None`.
    fn handle_command(&mut self, _cmd: &WidgetCommand) -> CommandValue {
        CommandValue::None
    }

    /// Get the tooltip text for this widget. Empty string means no tooltip.
    fn tooltip(&self) -> &str {
        ""
    }

    /// Set the tooltip text for this widget.
    fn set_tooltip(&mut self, _tooltip: &str) {}

    /// Get the anchor setting for this widget. Default: top-left (no resize).
    fn anchor(&self) -> Anchor {
        Anchor::TOP_LEFT
    }

    /// Set the anchor edges for this widget.
    fn set_anchor(&mut self, _anchor: Anchor) {}
}

/// Find a widget by name anywhere in a subtree, mutably — the mutable
/// counterpart of [`PanelWidget::find_rect`].
///
/// Free-standing rather than a trait method so it can hand back a
/// `&mut dyn PanelWidget` borrowed from the tree; a default method would
/// need `Self: Sized` and stop working on the trait objects containers hold.
pub fn find_widget_mut<'a>(
    root: &'a mut dyn PanelWidget,
    name: &str,
) -> Option<&'a mut dyn PanelWidget> {
    if root.name() == name {
        return Some(root);
    }
    for child in root.children_mut() {
        if let Some(found) = find_widget_mut(&mut **child, name) {
            return Some(found);
        }
    }
    None
}

/// Remove a widget by name from anywhere in a subtree and hand it back —
/// `removeChild` against a whole document rather than one parent.
pub fn take_widget(root: &mut dyn PanelWidget, name: &str) -> Option<Box<dyn PanelWidget>> {
    if let Some(w) = root.detach(name) {
        return Some(w);
    }
    for child in root.children_mut() {
        if let Some(w) = take_widget(child.as_mut(), name) {
            return Some(w);
        }
    }
    None
}

// ── NullWidget ─────────────────────────────────────────────────────────

/// A no-op widget used as a placeholder in containers.
pub struct NullWidget {
    rect: LayoutRect,
}

impl NullWidget {
    pub fn new() -> Self {
        Self {
            rect: LayoutRect::zero(),
        }
    }
}

impl PanelWidget for NullWidget {
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
    }
    fn rect(&self) -> LayoutRect {
        self.rect
    }
    fn render(&mut self, _ctx: &mut RenderContext) {}
    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool {
        false
    }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }
}

// ── FocusManager ───────────────────────────────────────────────────────

/// Manages keyboard focus across a collection of `PanelWidget`s.
///
/// Tracks which widget index is focused, supports Tab/Shift+Tab cycling,
/// and routes key events to the focused widget.
pub struct FocusManager {
    /// Index of the currently focused widget (into the caller's widget list).
    focused: Option<usize>,
    /// Index of the widget the mouse is currently hovering over.
    hovered: Option<usize>,
    /// Tooltip display state: (hover_start_frame, mouse_x, mouse_y).
    tooltip_state: Option<(u64, f32, f32)>,
    /// Frame counter incremented each time hover is updated.
    frame_counter: u64,
    /// Number of frames to wait before showing a tooltip (approx ~60fps → 45 frames ≈ 750ms).
    tooltip_delay_frames: u64,
}

impl FocusManager {
    pub fn new() -> Self {
        Self {
            focused: None,
            hovered: None,
            tooltip_state: None,
            frame_counter: 0,
            tooltip_delay_frames: 45,
        }
    }

    /// The currently focused widget index.
    pub fn focused(&self) -> Option<usize> {
        self.focused
    }

    /// Set the focused widget index directly.
    pub fn set_focused(&mut self, idx: Option<usize>) {
        self.focused = idx;
    }

    /// Focus the next focusable widget (Tab key).
    /// `count` is the total number of widgets.
    pub fn focus_next(&mut self, widgets: &mut [Box<dyn PanelWidget>]) {
        let len = widgets.len();
        if len == 0 {
            return;
        }
        let start = self.focused.map(|i| i + 1).unwrap_or(0);
        for offset in 0..len {
            let idx = (start + offset) % len;
            if widgets[idx].focusable() {
                self.apply_focus(widgets, Some(idx));
                return;
            }
        }
    }

    /// Focus the previous focusable widget (Shift+Tab key).
    pub fn focus_prev(&mut self, widgets: &mut [Box<dyn PanelWidget>]) {
        let len = widgets.len();
        if len == 0 {
            return;
        }
        let start = self.focused.unwrap_or(0).wrapping_sub(1);
        for offset in 0..len {
            let idx = (start.wrapping_sub(offset)) % len;
            if idx < len && widgets[idx].focusable() {
                self.apply_focus(widgets, Some(idx));
                return;
            }
        }
    }

    /// Focus the widget at the given position (e.g. on mouse click).
    pub fn focus_at(&mut self, widgets: &mut [Box<dyn PanelWidget>], x: f32, y: f32) {
        for (i, w) in widgets.iter().enumerate() {
            if w.focusable() && w.rect().contains(x, y) {
                self.apply_focus(widgets, Some(i));
                return;
            }
        }
        // Click on non-focusable area clears focus
        self.apply_focus(widgets, None);
    }

    /// Handle a key event: Tab/Shift+Tab for cycling, otherwise route to focused widget.
    /// Returns `true` if the event was consumed.
    pub fn handle_key(&mut self, widgets: &mut [Box<dyn PanelWidget>], event: &KeyEvent) -> bool {
        use winit::keyboard::Key::Named;
        use winit::keyboard::NamedKey;
        if event.state == winit::event::ElementState::Pressed {
            if let Named(NamedKey::Tab) = &event.logical_key {
                if event.shift {
                    self.focus_prev(widgets);
                } else {
                    self.focus_next(widgets);
                }
                return true;
            }
        }
        // Route to focused widget
        if let Some(idx) = self.focused {
            if idx < widgets.len() {
                return widgets[idx].handle_key(event);
            }
        }
        false
    }

    /// Handle a mouse event: update focus on click, update hover, then delegate.
    /// Returns `true` if the event was consumed.
    pub fn handle_mouse(
        &mut self,
        widgets: &mut [Box<dyn PanelWidget>],
        event: &MouseEvent,
    ) -> bool {
        if matches!(event.kind, MouseEventKind::Press(_)) {
            self.focus_at(widgets, event.x, event.y);
        }
        // Update hover tracking on every mouse event
        self.update_hover(widgets, event.x, event.y);
        // Delegate to all widgets (not just focused — for hover etc.)
        let mut consumed = false;
        for w in widgets.iter_mut() {
            if w.handle_mouse(event) {
                consumed = true;
            }
        }
        consumed
    }

    /// Update hover state: fire MouseEnter/MouseLeave when the hovered widget changes.
    pub fn update_hover(&mut self, widgets: &mut [Box<dyn PanelWidget>], x: f32, y: f32) {
        self.frame_counter += 1;
        let mut new_hover: Option<usize> = None;
        // Find topmost widget under cursor (last in list = topmost)
        for (i, w) in widgets.iter().enumerate().rev() {
            if w.rect().contains(x, y) {
                new_hover = Some(i);
                break;
            }
        }
        if new_hover != self.hovered {
            // Leave old
            if let Some(old) = self.hovered {
                if old < widgets.len() {
                    widgets[old].set_hovered(false);
                }
            }
            // Enter new
            if let Some(idx) = new_hover {
                if idx < widgets.len() {
                    widgets[idx].set_hovered(true);
                }
            }
            self.hovered = new_hover;
            // Reset tooltip timer on hover change
            self.tooltip_state = new_hover.map(|_| (self.frame_counter, x, y));
        }
    }

    /// Render all widgets, drawing a focus ring on the focused one.
    pub fn render_all(&self, widgets: &mut [Box<dyn PanelWidget>], ctx: &mut RenderContext) {
        for (i, w) in widgets.iter_mut().enumerate() {
            w.render(ctx);

            if self.focused == Some(i) {
                Self::draw_focus_ring(ctx, w.rect());
            }
        }
        // Render tooltip overlay if applicable
        self.render_tooltip(widgets, ctx);
    }

    /// Render the tooltip for the currently hovered widget after the delay has elapsed.
    fn render_tooltip(&self, widgets: &[Box<dyn PanelWidget>], ctx: &mut RenderContext) {
        let Some(idx) = self.hovered else {
            return;
        };
        let Some((start_frame, mx, my)) = self.tooltip_state else {
            return;
        };
        if self.frame_counter.saturating_sub(start_frame) < self.tooltip_delay_frames {
            return;
        }
        if idx >= widgets.len() {
            return;
        }
        let tip = widgets[idx].tooltip();
        if tip.is_empty() {
            return;
        }

        let font_size = 12.0_f32;
        let padding = 4.0_f32;
        let text_w = super::ide_text::measure_text(ctx.font_system, tip, font_size, ctx.scale);
        let tip_w = text_w + padding * 2.0;
        let tip_h = font_size + padding * 2.0;
        let tx = mx + 12.0;
        let ty = my + 18.0;

        let scale = ctx.scale;
        let ts = tiny_skia::Transform::from_scale(scale, scale);

        // Background
        if let Some(rect) =
            tiny_skia::Rect::from_xywh(tx * scale, ty * scale, tip_w * scale, tip_h * scale)
        {
            let mut bg = tiny_skia::Paint::default();
            bg.set_color_rgba8(255, 255, 225, 240); // light yellow tooltip bg
            ctx.pixmap
                .fill_rect(rect, &bg, tiny_skia::Transform::identity(), None);
            // Border
            let mut border = tiny_skia::Paint::default();
            border.set_color_rgba8(100, 100, 100, 200);
            let mut stroke = tiny_skia::Stroke::default();
            stroke.width = 1.0;
            if let Some(path) = super::rounded_rect_path(tx, ty, tip_w, tip_h, 2.0) {
                ctx.pixmap.stroke_path(&path, &border, &stroke, ts, None);
            }
        }

        // Text
        super::ide_text::draw_text(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            tip,
            tx + padding,
            ty + padding - 1.0,
            font_size,
            cosmic_text::Color::rgba(20, 20, 20, 255),
            scale,
        );
    }

    /// Drain events from all widgets into a single vec.
    /// Also enforces radio-group mutual exclusion: when a `RadioSelected` event
    /// is seen, other radios in the same group are deselected via commands.
    pub fn drain_all_events(&mut self, widgets: &mut [Box<dyn PanelWidget>]) -> Vec<WidgetEvent> {
        let mut all = Vec::new();
        for w in widgets.iter_mut() {
            all.append(&mut w.drain_events());
        }
        // Enforce radio group mutual exclusion
        for ev in &all {
            if let WidgetEvent::RadioSelected(selected_name, true) = ev {
                // Find the group of the selected radio
                let get_group = WidgetCommand::Custom("GetGroup".into(), CommandValue::None);
                let mut group = String::new();
                for w in widgets.iter_mut() {
                    if w.name() == selected_name.as_str() {
                        if let CommandValue::Text(g) = w.handle_command(&get_group) {
                            group = g;
                        }
                        break;
                    }
                }
                if !group.is_empty() {
                    // Deselect all other radios in the same group
                    for w in widgets.iter_mut() {
                        if w.name() != selected_name.as_str() {
                            if let CommandValue::Text(g) = w.handle_command(&get_group) {
                                if g == group {
                                    w.handle_command(&WidgetCommand::SetChecked(false));
                                }
                            }
                        }
                    }
                }
            }
        }
        all
    }

    /// Send a command to the first widget whose `name()` matches.
    /// Returns the command result, or `CommandValue::None` if no widget matched.
    pub fn send_command(
        &mut self,
        widgets: &mut [Box<dyn PanelWidget>],
        name: &str,
        cmd: &WidgetCommand,
    ) -> CommandValue {
        for w in widgets.iter_mut() {
            if let Some(result) = w.send_command_named(name, cmd) {
                return result;
            }
        }
        CommandValue::None
    }

    /// Send a command to the widget with the given `WidgetId`.
    /// Returns the command result, or `CommandValue::None` if no widget matched.
    pub fn send_command_by_id(
        &mut self,
        widgets: &mut [Box<dyn PanelWidget>],
        id: WidgetId,
        cmd: &WidgetCommand,
    ) -> CommandValue {
        for w in widgets.iter_mut() {
            if w.widget_id() == id {
                return w.handle_command(cmd);
            }
        }
        CommandValue::None
    }

    /// Broadcast a command to all widgets.
    /// Returns a `Vec` of `(WidgetId, CommandValue)` from widgets that returned a non-None value.
    pub fn broadcast_command(
        &mut self,
        widgets: &mut [Box<dyn PanelWidget>],
        cmd: &WidgetCommand,
    ) -> Vec<(WidgetId, CommandValue)> {
        let mut results = Vec::new();
        for w in widgets.iter_mut() {
            let val = w.handle_command(cmd);
            if !matches!(val, CommandValue::None) {
                results.push((w.widget_id(), val));
            }
        }
        results
    }

    // ── internal ──

    fn apply_focus(&mut self, widgets: &mut [Box<dyn PanelWidget>], new: Option<usize>) {
        if self.focused == new {
            return;
        }
        if let Some(old) = self.focused {
            if old < widgets.len() {
                widgets[old].set_focused(false);
            }
        }
        self.focused = new;
        if let Some(idx) = new {
            if idx < widgets.len() {
                widgets[idx].set_focused(true);
            }
        }
    }

    fn draw_focus_ring(ctx: &mut RenderContext, r: LayoutRect) {
        let scale = ctx.scale;
        let px = (r.x * scale) as i32;
        let py = (r.y * scale) as i32;
        let pw = (r.w * scale) as i32;
        let ph = (r.h * scale) as i32;
        let color = ColorU8::from_rgba(0, 120, 215, 200).premultiply();
        let thickness = (2.0 * scale) as i32;
        let pix = &mut *ctx.pixmap;
        let w = pix.width() as i32;
        let h = pix.height() as i32;
        // top edge
        for dy in 0..thickness {
            let y = py + dy;
            if y < 0 || y >= h {
                continue;
            }
            for x in px..(px + pw).min(w) {
                if x >= 0 {
                    pix.pixels_mut()[(y * w + x) as usize] = color;
                }
            }
        }
        // bottom edge
        for dy in 0..thickness {
            let y = py + ph - 1 - dy;
            if y < 0 || y >= h {
                continue;
            }
            for x in px..(px + pw).min(w) {
                if x >= 0 {
                    pix.pixels_mut()[(y * w + x) as usize] = color;
                }
            }
        }
        // left edge
        for dx in 0..thickness {
            let x = px + dx;
            if x < 0 || x >= w {
                continue;
            }
            for y in py..(py + ph).min(h) {
                if y >= 0 {
                    pix.pixels_mut()[(y * w + x) as usize] = color;
                }
            }
        }
        // right edge
        for dx in 0..thickness {
            let x = px + pw - 1 - dx;
            if x < 0 || x >= w {
                continue;
            }
            for y in py..(py + ph).min(h) {
                if y >= 0 {
                    pix.pixels_mut()[(y * w + x) as usize] = color;
                }
            }
        }
    }
}
