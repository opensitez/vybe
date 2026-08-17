//! CSS declarations: storage, parsing, and the typed view layout reads.
//!
//! Lifted and adapted from the `osz-htmledit` HTML editor's `css.rs`. That
//! project renders through GTK text tags; this one renders through our own
//! widgets, so only the toolkit-independent half came across — the declaration
//! parser, the length/shorthand rules, and the property record. None of the
//! GTK application code did.
//!
//! ## Two layers, deliberately
//!
//! 1. [`Style`] is the **store**: the declarations exactly as they were set,
//!    keyed by property name. It accepts anything, including properties nothing
//!    renders. That is what makes `el.style.color = 'red'` read back as `'red'`
//!    — before this, a style write was translated straight into a widget command
//!    and the CSS was forgotten, so the read side could only answer for geometry
//!    it could recover from the widget's rect.
//! 2. [`CssProperties`] is the **typed view**: the subset layout and painting
//!    act on, parsed into enums and lengths. Unknown properties never reach it
//!    and do not need to — they are still stored, still serialise, still round-
//!    trip.
//!
//! This is the CSSOM/layout split: the object model records what was said, and
//! layout consumes what it understands. A property being unimplemented is then a
//! rendering gap rather than data loss, which is the difference between a
//! `display: grid` that does nothing yet and one that silently disappears.
//!
//! ## What is deliberately absent
//!
//! **Selectors and the cascade.** Every frontend sets style through the CSSOM
//! (`element.style.setProperty`); a compiled VCL, WinForms or Flutter program
//! contains no stylesheet, so there is nothing to match against. The source this
//! came from resolved a fixed four-tier order (tag < .class < tag.class < #id <
//! inline) with descendant selectors flattened to their last component and every
//! combinator dropped — adequate for an editor, but not a cascade, and it would
//! answer wrongly on a real stylesheet. It was left behind rather than renamed.

use std::collections::BTreeMap;

// ── Values ──────────────────────────────────────────────────────────────────

/// A CSS length.
///
/// Percentages stay **symbolic**. The source this came from resolved `%` to
/// `value * 0.16` ("100% ≈ 16px base") at parse time, which is invisible in a
/// text editor and wrong anywhere with a containing block: `width: 50%` is a
/// fraction of a parent that parsing cannot see. Resolution belongs to layout,
/// so it is deferred to layout.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Length {
    Px(f32),
    Percent(f32),
    Auto,
}

/// Where an item sits on one grid axis — `grid-column` / `grid-row`.
///
/// Three forms, because CSS has three and they are genuinely different
/// questions: `auto` lets the placement cursor decide, a LINE pins the item to
/// a numbered grid line, and a SPAN says how many tracks it covers without
/// saying where it starts.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum GridLine {
    Auto,
    /// A 1-based grid line, as CSS numbers them.
    Line(i32),
    Span(u32),
}

impl GridLine {
    pub fn parse(token: &str) -> Option<GridLine> {
        let token = token.trim();
        if token.is_empty() || token.eq_ignore_ascii_case("auto") {
            return Some(GridLine::Auto);
        }
        if let Some(n) = token.to_ascii_lowercase().strip_prefix("span") {
            return n.trim().parse::<u32>().ok().map(GridLine::Span);
        }
        token.parse::<i32>().ok().map(GridLine::Line)
    }

    pub fn as_css(self) -> String {
        match self {
            GridLine::Auto => "auto".to_string(),
            GridLine::Line(n) => n.to_string(),
            GridLine::Span(n) => format!("span {n}"),
        }
    }
}

/// **A track template, with the auto-repeating section kept separate.**
///
/// `repeat(4, 1fr)` expands at parse time — the count is written down. But
/// `repeat(auto-fill, minmax(200px, 1fr))` cannot: how many times it repeats
/// depends on how much room the container has, which parsing does not know and
/// layout does. So the pattern is CARRIED, not expanded, and resolved against a
/// real extent by [`TrackTemplate::resolve`].
///
/// `before`/`after` are the fixed tracks either side of it, because
/// `200px repeat(auto-fill, 1fr) 200px` is legal and the sidebars must keep
/// their places.
#[derive(Clone, Debug, PartialEq, Default)]
pub struct TrackTemplate {
    pub before: Vec<TrackSize>,
    pub auto_repeat: Option<(AutoRepeat, Vec<TrackSize>)>,
    pub after: Vec<TrackSize>,
}

/// `auto-fill` vs `auto-fit` — they place tracks identically and differ only in
/// what happens to the EMPTY ones: `auto-fit` collapses them, so the filled
/// tracks absorb the space.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AutoRepeat {
    Fill,
    Fit,
}

impl TrackTemplate {
    pub fn is_empty(&self) -> bool {
        self.before.is_empty() && self.auto_repeat.is_none() && self.after.is_empty()
    }

    /// The concrete track list for a container of this size.
    ///
    /// The repetition count is the largest number of whole patterns that fit —
    /// Grid §7.2.3.1 — computed from each pattern track's own floor, because
    /// that is the smallest it can ever be and therefore the most that can fit.
    /// A pattern with no definite floor (`1fr` alone) repeats once: it would
    /// otherwise divide by zero and fill forever.
    pub fn resolve(&self, extent: f32, gap: f32) -> Vec<TrackSize> {
        let mut out = self.before.clone();
        if let Some((_, pattern)) = &self.auto_repeat {
            let floor: f32 = pattern.iter().map(|t| t.min_extent(extent)).sum();
            let fixed: f32 = self
                .before
                .iter()
                .chain(self.after.iter())
                .map(|t| t.min_extent(extent))
                .sum();
            let leading = self.before.len() + self.after.len();
            let count = if floor > 0.0 {
                // Each repetition costs its floor plus the gap that precedes it.
                let room = extent - fixed - gap * leading.saturating_sub(1) as f32;
                (((room + gap) / (floor + gap * pattern.len() as f32)).floor() as i64).max(1)
            } else {
                1
            };
            for _ in 0..count.min(1024) {
                out.extend(pattern.iter().copied());
            }
        }
        out.extend(self.after.iter().copied());
        if out.is_empty() {
            out.push(TrackSize::Fr(1.0));
        }
        out
    }
}

/// `grid-template-areas: "head head" "side main"` — one quoted string per row.
///
/// Every row must name the same number of cells; a ragged template is invalid
/// and dropped whole, because a half-applied grid places items in cells the
/// author never wrote.
pub fn parse_area_template(value: &str) -> Option<Vec<Vec<String>>> {
    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut rest = value.trim();
    while let Some(open) = rest.find(['"', '\'']) {
        let quote = rest.as_bytes()[open] as char;
        let after = &rest[open + 1..];
        let close = after.find(quote)?;
        let cells: Vec<String> = after[..close]
            .split_whitespace()
            .map(str::to_string)
            .collect();
        if cells.is_empty() {
            return None;
        }
        rows.push(cells);
        rest = after[close + 1..].trim_start();
    }
    if rows.is_empty() {
        return None;
    }
    let width = rows[0].len();
    rows.iter().all(|r| r.len() == width).then_some(rows)
}

/// Serialise an area template back to CSS, so the cascade diff can carry it.
pub fn area_template_css(rows: &[Vec<String>]) -> String {
    rows.iter()
        .map(|r| format!("\"{}\"", r.join(" ")))
        .collect::<Vec<_>>()
        .join(" ")
}

/// `grid-column: <start> / <end>` — the shorthand, which is why it answers two
/// values. A missing `/` means the start alone, with the end left `auto`.
pub fn parse_grid_placement(value: &str) -> (Option<GridLine>, Option<GridLine>) {
    match value.split_once('/') {
        Some((start, end)) => (GridLine::parse(start), GridLine::parse(end)),
        None => (GridLine::parse(value), None),
    }
}

/// One track of a grid template — a column width or a row height.
///
/// A [`Length`] cannot spell this: `fr` is a **fraction of the leftover**, so
/// it is not a length at all until every fixed track has been subtracted. That
/// is the whole reason grid needs its own size type rather than reusing the box
/// one.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum TrackSize {
    Px(f32),
    Percent(f32),
    /// `1fr` — a share of what is left after the fixed tracks and the gaps.
    Fr(f32),
    /// `auto` — as big as the largest item in the track.
    Auto,
    /// **`minmax(min, max)`** — a floor and a ceiling on one track.
    ///
    /// Not expressible as a single size, which is why `TrackSize` needed a
    /// second shape rather than another scalar: `minmax(200px, 1fr)` is a track
    /// that never shrinks below 200px and otherwise takes a share of the
    /// leftover, and neither half alone says that.
    ///
    /// `fr` is invalid in the MIN position per Grid §7.2.2 — a floor cannot be
    /// a share of what is left over, because the leftover is not known until
    /// the floors are. Parsed as `auto` there rather than rejected, which is
    /// what a browser's error handling amounts to for this case.
    MinMax(MinMaxSide, MinMaxSide),
}

/// One half of a `minmax()`. A separate type so `TrackSize` stays `Copy` —
/// a boxed recursive track would make every track list an allocation.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum MinMaxSide {
    Px(f32),
    Percent(f32),
    Fr(f32),
    Auto,
}

impl MinMaxSide {
    fn parse(token: &str) -> Option<MinMaxSide> {
        match TrackSize::parse_basic(token)? {
            TrackSize::Px(v) => Some(MinMaxSide::Px(v)),
            TrackSize::Percent(p) => Some(MinMaxSide::Percent(p)),
            TrackSize::Fr(f) => Some(MinMaxSide::Fr(f)),
            TrackSize::Auto => Some(MinMaxSide::Auto),
            TrackSize::MinMax(..) => None,
        }
    }

    fn as_css(self) -> String {
        match self {
            MinMaxSide::Px(v) => format!("{v}px"),
            MinMaxSide::Percent(p) => format!("{p}%"),
            MinMaxSide::Fr(f) => format!("{f}fr"),
            MinMaxSide::Auto => "auto".to_string(),
        }
    }

    /// The side as a definite length, when it is one. `fr` and `auto` answer
    /// `None` — both need something layout knows and parsing does not.
    pub fn definite(self, extent: f32) -> Option<f32> {
        match self {
            MinMaxSide::Px(v) => Some(v),
            MinMaxSide::Percent(p) => Some(extent * p / 100.0),
            _ => None,
        }
    }

    pub fn fr(self) -> Option<f32> {
        match self {
            MinMaxSide::Fr(f) => Some(f),
            _ => None,
        }
    }
}

impl TrackSize {
    /// A track size that is NOT a `minmax()` — the four scalar forms.
    ///
    /// `em`/`rem`/`%`/`px` all arrive through [`parse_length`], so a track list
    /// accepts every length unit the rest of the engine does; only `fr` is
    /// grid's own and handled here.
    fn parse_basic(token: &str) -> Option<TrackSize> {
        let token = token.trim();
        let lower = token.to_ascii_lowercase();
        if let Some(n) = lower.strip_suffix("fr") {
            return n.trim().parse::<f32>().ok().map(TrackSize::Fr);
        }
        match parse_length(token)? {
            Length::Px(v) => Some(TrackSize::Px(v)),
            Length::Percent(p) => Some(TrackSize::Percent(p)),
            Length::Auto => Some(TrackSize::Auto),
        }
    }

    pub fn parse(token: &str) -> Option<TrackSize> {
        let token = token.trim();
        let lower = token.to_ascii_lowercase();
        if let Some(rest) = lower.strip_prefix("minmax(") {
            let close = matching_paren(rest)?;
            // Split on the ORIGINAL casing so a unit keyword is not folded, and
            // on the top-level comma only.
            let inner = &token["minmax(".len().."minmax(".len() + close];
            let (min, max) = inner.split_once(',')?;
            let min = MinMaxSide::parse(min)?;
            // A floor cannot be a share of the leftover — §7.2.2.
            let min = if min.fr().is_some() { MinMaxSide::Auto } else { min };
            return Some(TrackSize::MinMax(min, MinMaxSide::parse(max)?));
        }
        Self::parse_basic(token)
    }

    pub fn as_css(self) -> String {
        match self {
            TrackSize::Px(v) => format!("{v}px"),
            TrackSize::Percent(p) => format!("{p}%"),
            TrackSize::Fr(f) => format!("{f}fr"),
            TrackSize::Auto => "auto".to_string(),
            TrackSize::MinMax(a, b) => format!("minmax({}, {})", a.as_css(), b.as_css()),
        }
    }

    /// The smallest this track can be, as a definite length — its FLOOR.
    ///
    /// What `auto-fill` counts with: the floor is the most that can fit, so it
    /// is what decides how many repetitions there is room for. `fr` and `auto`
    /// have no floor and answer `0`, which is why a pattern made only of them
    /// repeats once instead of forever.
    pub fn min_extent(self, extent: f32) -> f32 {
        match self {
            TrackSize::Px(v) => v.max(0.0),
            TrackSize::Percent(p) => (extent * p / 100.0).max(0.0),
            TrackSize::MinMax(min, _) => min.definite(extent).unwrap_or(0.0).max(0.0),
            TrackSize::Fr(_) | TrackSize::Auto => 0.0,
        }
    }
}

/// Serialise a track list back to CSS, so the cascade diff can carry it to the
/// widget as a string like every other declaration.
pub fn track_list_css(tracks: &[TrackSize]) -> String {
    tracks
        .iter()
        .map(|t| t.as_css())
        .collect::<Vec<_>>()
        .join(" ")
}

/// `grid-template-columns` / `grid-template-rows`.
///
/// Handles `repeat(n, …)`, including nested lists, because a template written
/// any other way is unreadable past three columns. An unparseable token drops
/// the whole declaration rather than silently shortening the grid — a template
/// with a missing track would place every item in the wrong cell.
/// Parse a full template, keeping any `repeat(auto-fill|auto-fit, …)` whole.
///
/// Split rather than expanded, because the count is not knowable here — see
/// [`TrackTemplate`].
pub fn parse_track_template(input: &str) -> Option<TrackTemplate> {
    let lower = input.to_ascii_lowercase();
    let Some(at) = lower.find("repeat(") else {
        return parse_track_list(input).map(|tracks| TrackTemplate {
            before: tracks,
            ..TrackTemplate::default()
        });
    };
    let after_open = &input[at + "repeat(".len()..];
    let close = matching_paren(after_open)?;
    let inner = &after_open[..close];
    let (count, pattern) = inner.split_once(',')?;
    let kind = match count.trim().to_ascii_lowercase().as_str() {
        "auto-fill" => AutoRepeat::Fill,
        "auto-fit" => AutoRepeat::Fit,
        // A numeric `repeat()` is expandable here and always was.
        _ => {
            return parse_track_list(input).map(|tracks| TrackTemplate {
                before: tracks,
                ..TrackTemplate::default()
            });
        }
    };
    let before = input[..at].trim();
    let after = after_open[close + 1..].trim();
    Some(TrackTemplate {
        before: if before.is_empty() {
            Vec::new()
        } else {
            parse_track_list(before)?
        },
        auto_repeat: Some((kind, parse_track_list(pattern)?)),
        after: if after.is_empty() {
            Vec::new()
        } else {
            parse_track_list(after)?
        },
    })
}

/// Serialise a template, keeping an auto-repeat in its unexpanded form so a
/// round trip through the cascade does not freeze the count.
pub fn track_template_css(t: &TrackTemplate) -> String {
    let mut parts: Vec<String> = t.before.iter().map(|x| x.as_css()).collect();
    if let Some((kind, pattern)) = &t.auto_repeat {
        let word = match kind {
            AutoRepeat::Fill => "auto-fill",
            AutoRepeat::Fit => "auto-fit",
        };
        parts.push(format!("repeat({word}, {})", track_list_css(pattern)));
    }
    parts.extend(t.after.iter().map(|x| x.as_css()));
    parts.join(" ")
}

pub fn parse_track_list(input: &str) -> Option<Vec<TrackSize>> {
    let mut out: Vec<TrackSize> = Vec::new();
    let mut rest = input.trim();
    while !rest.is_empty() {
        let lower = rest.to_ascii_lowercase();
        if lower.starts_with("repeat(") {
            let after = &rest["repeat(".len()..];
            let close = matching_paren(after)?;
            let (count, list) = after[..close].split_once(',')?;
            let count: usize = count.trim().parse().ok()?;
            let tracks = parse_track_list(list)?;
            // A bound, because `repeat(100000, 1fr)` is a denial of service
            // rather than a layout. CSS has no such limit; a renderer does.
            if count * tracks.len() > 1024 {
                return None;
            }
            for _ in 0..count {
                out.extend(tracks.iter().copied());
            }
            rest = after[close + 1..].trim_start();
            continue;
        }
        // A token ends at whitespace — but NOT whitespace inside parentheses.
        // `minmax(200px, 1fr)` is one track and splitting it on the space after
        // the comma would make it two unparseable ones.
        let mut depth = 0usize;
        let mut end = rest.len();
        for (i, ch) in rest.char_indices() {
            match ch {
                '(' => depth += 1,
                ')' => depth = depth.saturating_sub(1),
                c if c.is_whitespace() && depth == 0 => {
                    end = i;
                    break;
                }
                _ => {}
            }
        }
        let (token, remainder) = rest.split_at(end);
        out.push(TrackSize::parse(token)?);
        rest = remainder.trim_start();
    }
    (!out.is_empty()).then_some(out)
}

/// Index of the `)` closing a `(` that has already been consumed.
fn matching_paren(input: &str) -> Option<usize> {
    let mut depth = 0usize;
    for (i, ch) in input.char_indices() {
        match ch {
            '(' => depth += 1,
            ')' if depth == 0 => return Some(i),
            ')' => depth -= 1,
            _ => {}
        }
    }
    None
}

impl Length {
    /// The pixel value, when it is one. `Percent` and `Auto` answer `None` —
    /// both need a containing block.
    pub fn px(self) -> Option<f32> {
        match self {
            Length::Px(v) => Some(v),
            _ => None,
        }
    }

    /// Resolve against a containing-block extent.
    pub fn resolve(self, basis: f32) -> Option<f32> {
        match self {
            Length::Px(v) => Some(v),
            Length::Percent(p) => Some(basis * p / 100.0),
            Length::Auto => None,
        }
    }
}

impl std::fmt::Display for Length {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Length::Px(v) => write!(f, "{v}px"),
            Length::Percent(v) => write!(f, "{v}%"),
            Length::Auto => write!(f, "auto"),
        }
    }
}

macro_rules! keyword_enum {
    ($(#[$m:meta])* $name:ident { $($variant:ident => $css:literal),+ $(,)? }) => {
        $(#[$m])*
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        pub enum $name { $($variant),+ }

        impl $name {
            pub fn parse(value: &str) -> Option<Self> {
                match value.trim().to_ascii_lowercase().as_str() {
                    $($css => Some($name::$variant),)+
                    _ => None,
                }
            }

            pub fn as_css(self) -> &'static str {
                match self { $($name::$variant => $css),+ }
            }
        }
    };
}

keyword_enum! {
    /// The box's own layout mode.
    ///
    /// `none` is NOT here — it is visibility, and it is recorded separately
    /// (`display_none`). The two were conflated once: the whole `display`
    /// property meant "is it visible", so `display: flex` marked an element
    /// visible and then did nothing, which reads as an unimplemented feature
    /// rather than a consumed one.
    /// **`inline-flex`/`inline-grid` are inline-level OUTSIDE and
    /// flex/grid INSIDE.** `display` names two things — how the box takes part
    /// in its PARENT's flow, and what formatting context it establishes for its
    /// CHILDREN — and these are the values where the two answers differ. The
    /// engine already models exactly that split: the box is told its context,
    /// the parent is told what kind of box arrived.
    /// **The table display types** — CSS 2.1 §17.2.
    ///
    /// A table is not a grid with nicer names. Its columns are sized from what
    /// the CELLS contain (§17.5.2), which no other formatting context does:
    /// grid tracks come from a template the author wrote, flex bases from the
    /// items. That is why `<table>` cannot simply be `display: grid`, and why
    /// mapping it to the `datagridview` WIDGET was worse still — it made an
    /// HTML layout impersonate a .NET control.
    ///
    /// `table-column`/`table-column-group` generate NO box. They exist to carry
    /// a width down onto a column, which is the only thing `<col>` does.
    Display {
        Block => "block",
        Flex => "flex",
        InlineBlock => "inline-block",
        Inline => "inline",
        Grid => "grid",
        InlineFlex => "inline-flex",
        InlineGrid => "inline-grid",
        Table => "table",
        InlineTable => "inline-table",
        TableRow => "table-row",
        TableRowGroup => "table-row-group",
        TableHeaderGroup => "table-header-group",
        TableFooterGroup => "table-footer-group",
        TableCell => "table-cell",
        TableCaption => "table-caption",
        TableColumn => "table-column",
        TableColumnGroup => "table-column-group",
    }
}

keyword_enum! {
    /// `border-collapse` — CSS 2.1 §17.6.
    ///
    /// Which of the two border models a table uses. `separate` gives every cell
    /// its own border with `border-spacing` between them; `collapse` merges
    /// adjacent borders into one, and the winner is decided by a conflict
    /// resolution order (§17.6.2), not by whoever drew last.
    BorderCollapse {
        Separate => "separate",
        Collapse => "collapse",
    }
}

keyword_enum! {
    /// `table-layout` — CSS 2.1 §17.5.2.
    ///
    /// **The two column-width algorithms.** `auto` measures every cell's
    /// content, so the table cannot be laid out until all of it is known;
    /// `fixed` takes the widths from the columns and the first row alone and
    /// never looks at the rest, which is why it is the fast one.
    TableLayout {
        Auto => "auto",
        Fixed => "fixed",
    }
}

keyword_enum! {
    /// **Take this box out of flow and shift it to one side** — CSS 2.1 §9.5.
    ///
    /// The oldest layout mode on the web and the one this engine never had.
    /// A float is not merely "positioned left": it is removed from the flow,
    /// pushed until its margin edge meets the containing block or an earlier
    /// float, and then — the part that makes it a MODE rather than an offset —
    /// the LINE BOXES beside it are shortened so text wraps around it.
    Float {
        None => "none",
        Left => "left",
        Right => "right",
    }
}

keyword_enum! {
    /// **Move below the floats already placed** — §9.5.2.
    ///
    /// The other half of the float model, and the half a naive implementation
    /// forgets: without `clear`, a short paragraph beside a tall image can
    /// never be made to start under it, and every "why is my footer inside the
    /// sidebar" bug is this.
    Clear {
        None => "none",
        Left => "left",
        Right => "right",
        Both => "both",
    }
}

keyword_enum! {
    /// Whether text WRAPS, and what happens to the source's own whitespace.
    ///
    /// The half that matters here is wrapping: the shaper already takes an
    /// optional wrap width, and `nowrap` is that width being absent. Without
    /// this, a label in a narrow box always broke, and no frontend could say
    /// otherwise — a toolbar caption or a table cell that must stay on one line
    /// had no way to ask.
    ///
    /// ⚠ The whitespace-PROCESSING half (`pre` keeping runs of spaces and
    /// newlines) is parsed and not yet honoured: collapsing happens when text
    /// enters the DOM, so `pre` has to change that, not the paint. `Pre` and
    /// `PreWrap` therefore only carry their wrapping behaviour today, which is
    /// recorded rather than hidden.
    WhiteSpace {
        Normal => "normal",
        Nowrap => "nowrap",
        Pre => "pre",
        PreWrap => "pre-wrap",
        PreLine => "pre-line",
    }
}

keyword_enum! {
    /// How a box is positioned relative to its container.
    ///
    /// `Absolute` is what every pixel-positioned frontend means by setting
    /// `Left`/`Top`: out of flow, coordinates honoured, the container does not
    /// get to rearrange it.
    Position {
        Static => "static",
        Relative => "relative",
        Absolute => "absolute",
        Fixed => "fixed",
        Sticky => "sticky",
    }
}

keyword_enum! {
    FlexDirection {
        Row => "row",
        RowReverse => "row-reverse",
        Column => "column",
        ColumnReverse => "column-reverse",
    }
}

keyword_enum! {
    FlexWrap {
        NoWrap => "nowrap",
        Wrap => "wrap",
        WrapReverse => "wrap-reverse",
    }
}

keyword_enum! {
    JustifyContent {
        FlexStart => "flex-start",
        FlexEnd => "flex-end",
        Center => "center",
        SpaceBetween => "space-between",
        SpaceAround => "space-around",
        SpaceEvenly => "space-evenly",
    }
}

keyword_enum! {
    /// `grid-auto-flow` — which axis auto-placement fills, and how hard it
    /// tries. `Dense` back-fills earlier holes; the sparse default never moves
    /// the cursor backwards.
    GridAutoFlow {
        Row => "row",
        Column => "column",
        RowDense => "row dense",
        ColumnDense => "column dense",
    }
}

keyword_enum! {
    AlignItems {
        FlexStart => "flex-start",
        FlexEnd => "flex-end",
        Center => "center",
        Baseline => "baseline",
        Stretch => "stretch",
    }
}

keyword_enum! {
    /// What a declared `width` and `height` MEASURE.
    ///
    /// The initial value is CSS's: `content-box`, so `width: 100px` plus
    /// `padding: 10px` is a 120px box. The toolkits this compiler targets mean
    /// the opposite — a VCL or WinForms control's `Width` includes its border
    /// and padding — but that is a statement those frontends make about their
    /// own controls, not a property of the box model, and it is declared as
    /// `box-sizing: border-box` on the elements it is true of (see
    /// [`crate::ua`]'s control rule). A frontend with different conventions,
    /// Flutter included, declares its own and gets it.
    BoxSizing {
        BorderBox => "border-box",
        ContentBox => "content-box",
    }
}

keyword_enum! {
    TextAlign {
        Left => "left",
        Right => "right",
        Center => "center",
        Justify => "justify",
    }
}

keyword_enum! {
    FontStyle {
        Normal => "normal",
        Italic => "italic",
        Oblique => "oblique",
    }
}

keyword_enum! {
    BorderStyle {
        None => "none",
        Solid => "solid",
        Dashed => "dashed",
        Dotted => "dotted",
        Double => "double",
        Hidden => "hidden",
    }
}

keyword_enum! {
    Overflow {
        Visible => "visible",
        Hidden => "hidden",
        Scroll => "scroll",
        Auto => "auto",
    }
}

/// Per-side box values, in CSS order.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Sides<T: Copy> {
    pub top: Option<T>,
    pub right: Option<T>,
    pub bottom: Option<T>,
    pub left: Option<T>,
}

// Hand-written: deriving `Default` would demand `T: Default`, and none of the
// side types have a meaningful default — an unspecified border style is absent,
// not `none`.
impl<T: Copy> Default for Sides<T> {
    fn default() -> Self {
        Self {
            top: None,
            right: None,
            bottom: None,
            left: None,
        }
    }
}

impl<T: Copy> Sides<T> {
    fn merge_from(&mut self, other: &Sides<T>) {
        if other.top.is_some() {
            self.top = other.top;
        }
        if other.right.is_some() {
            self.right = other.right;
        }
        if other.bottom.is_some() {
            self.bottom = other.bottom;
        }
        if other.left.is_some() {
            self.left = other.left;
        }
    }

    fn set_all(&mut self, value: T) {
        self.top = Some(value);
        self.right = Some(value);
        self.bottom = Some(value);
        self.left = Some(value);
    }
}

// ── The typed view ──────────────────────────────────────────────────────────

keyword_enum! {
    /// `cursor` — CSS UI §8.1, the subset a desktop toolkit can actually show.
    ///
    /// Inherited, which is what makes `pointer` on an `<a>` cover the `<em>`
    /// inside it without a second rule.
    ///
    /// This is the property that lets a LINK look clickable without the painter
    /// knowing what a link is. A run carries its resolved cursor exactly as it
    /// carries its resolved colour, so hard-coding "text runs are hands" — which
    /// would make every `<span>` a hand — is never needed.
    Cursor {
        Auto => "auto",
        Default => "default",
        Pointer => "pointer",
        Text => "text",
        Move => "move",
        NotAllowed => "not-allowed",
        Wait => "wait",
        Help => "help",
        Crosshair => "crosshair",
        Grab => "grab",
    }
}

/// `text-transform` — CSS Text §2.1.
///
/// A **rendering** transform, not a content one. `textContent` keeps what the
/// author wrote and only the painted run is re-cased, which is why this is
/// applied where inline runs are built rather than where text is stored: a
/// stylesheet that uppercases a label must not change what the DOM answers
/// when the program reads the label back.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextTransform {
    None,
    Uppercase,
    Lowercase,
    Capitalize,
}

impl TextTransform {
    pub fn parse(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "none" => Some(Self::None),
            "uppercase" => Some(Self::Uppercase),
            "lowercase" => Some(Self::Lowercase),
            "capitalize" => Some(Self::Capitalize),
            _ => None,
        }
    }

    pub fn as_css(&self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Uppercase => "uppercase",
            Self::Lowercase => "lowercase",
            Self::Capitalize => "capitalize",
        }
    }

    /// Apply to one run of text.
    ///
    /// `capitalize` titlecases the first letter of each WORD, and a word starts
    /// after any whitespace — so `"hello world"` becomes `"Hello World"` while
    /// `"o'neill"` keeps its inner letters, which is what the spec's "first
    /// typographic letter unit" means for the Latin text a form contains.
    /// Casing is Unicode-aware (`to_uppercase`, not `to_ascii_uppercase`), so
    /// `"ß"` and accented letters transform rather than passing through.
    pub fn apply(&self, text: &str) -> String {
        match self {
            Self::None => text.to_string(),
            Self::Uppercase => text.to_uppercase(),
            Self::Lowercase => text.to_lowercase(),
            Self::Capitalize => {
                let mut out = String::with_capacity(text.len());
                let mut at_word_start = true;
                for ch in text.chars() {
                    if at_word_start && ch.is_alphanumeric() {
                        out.extend(ch.to_uppercase());
                        at_word_start = false;
                    } else {
                        out.push(ch);
                        if ch.is_whitespace() {
                            at_word_start = true;
                        }
                    }
                }
                out
            }
        }
    }
}

/// The properties layout and painting act on, parsed.
///
/// Every field is optional: `None` is "not specified here", which is what lets
/// [`CssProperties::merge`] layer one set of declarations over another without a
/// specified/unspecified flag beside each value.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CssProperties {
    // Layout mode
    pub display: Option<Display>,
    /// `display: none` — visibility, not a layout mode. Kept apart so a real
    /// mode can be set without the two fighting.
    pub display_none: bool,
    pub position: Option<Position>,
    pub offsets: Sides<Length>,
    pub z_index: Option<i32>,
    pub overflow: Option<Overflow>,

    // Flex container
    pub flex_direction: Option<FlexDirection>,
    pub flex_wrap: Option<FlexWrap>,
    pub justify_content: Option<JustifyContent>,
    pub align_items: Option<AlignItems>,
    /// `align-content` — how WRAPPED LINES are distributed along the cross
    /// axis. Meaningless until lines can wrap, which is why it did not exist
    /// before `flex-wrap` was implemented. Flutter reaches it as
    /// `Wrap.runAlignment`.
    pub align_content: Option<JustifyContent>,
    pub gap: Option<Length>,
    /// The two gap axes as their own fields.
    ///
    /// `gap` alone cannot carry them: all three spellings used to parse into
    /// the single field above, so a rule saying `column-gap: 20px; row-gap:
    /// 10px` reached the cascade as ONE value and both axes got whichever
    /// declaration came last. The shorthand still fills all three, which is
    /// what makes `gap: 8px` keep working unchanged.
    pub row_gap: Option<Length>,
    pub column_gap: Option<Length>,

    // Table container — CSS 2.1 §17. These live on the TABLE box, not on the
    // cells: which border model is in force and how wide the gaps are is one
    // decision for the whole table, which is why `border-collapse` inherits.
    pub border_collapse: Option<BorderCollapse>,
    /// `border-spacing` — the gap between cell borders in the separate model.
    /// One length, not two: the horizontal and vertical values are the same
    /// number in every table anyone writes, and a second field that nothing
    /// sets is a field that drifts.
    pub border_spacing: Option<Length>,
    pub table_layout: Option<TableLayout>,

    // Grid container. Two track lists and nothing else yet: auto-placement
    // fills them row-major, which is what a template alone means in CSS.
    pub grid_template_columns: Option<TrackTemplate>,
    pub grid_template_rows: Option<TrackTemplate>,
    /// `grid-template-areas` — one `Vec<String>` per row, one name per cell.
    ///
    /// The names ARE the layout: a cell holding `header` in two adjacent
    /// columns is one area two tracks wide, so the rectangle is derived rather
    /// than declared. `.` is an empty cell and is kept as-is, because "no area"
    /// is a name a lookup must miss rather than an absence to skip.
    pub grid_template_areas: Option<Vec<Vec<String>>>,
    pub grid_auto_flow: Option<GridAutoFlow>,
    /// `grid-auto-rows` / `grid-auto-columns` — the size of an IMPLICIT track,
    /// one the template never named. `auto` until told otherwise, which is why
    /// a grid with more items than cells grew content-sized rows before this.
    pub grid_auto_rows: Option<TrackSize>,
    pub grid_auto_columns: Option<TrackSize>,
    /// `justify-items` / `justify-self` — the INLINE-axis counterpart of
    /// `align-items`/`align-self`. Grid has both axes; flex only ever needed
    /// one, which is why these did not exist.
    pub justify_items: Option<AlignItems>,
    pub justify_self: Option<AlignItems>,

    // Grid item — where this box sits in its parent's grid.
    /// `grid-area: header` — the NAME form. The four-line form
    /// (`grid-area: 1 / 1 / 3 / 2`) writes the four longhands instead.
    pub grid_area: Option<String>,
    pub grid_column_start: Option<GridLine>,
    pub grid_column_end: Option<GridLine>,
    pub grid_row_start: Option<GridLine>,
    pub grid_row_end: Option<GridLine>,

    // Flex item
    pub flex_grow: Option<f32>,
    pub flex_shrink: Option<f32>,
    pub flex_basis: Option<Length>,
    pub align_self: Option<AlignItems>,
    pub order: Option<i32>,

    // Box
    pub width: Option<Length>,
    pub height: Option<Length>,
    pub min_width: Option<Length>,
    pub min_height: Option<Length>,
    pub max_width: Option<Length>,
    pub max_height: Option<Length>,
    pub margin: Sides<Length>,
    pub padding: Sides<Length>,
    pub border_width: Sides<f32>,
    pub border_style: Sides<BorderStyle>,
    pub border_color: Sides<u32>,
    pub border_radius: Option<f32>,
    /// What `width`/`height` measure. Unset means the CSS initial value,
    /// `content-box` — see [`BoxSizing`].
    pub box_sizing: Option<BoxSizing>,

    // Paint
    pub color: Option<u32>,
    pub background_color: Option<u32>,
    pub opacity: Option<f32>,
    pub visibility_hidden: bool,

    // Text
    pub font_family: Option<String>,
    /// In **pixels**. The source converted to points (`px / 1.333`) because
    /// Pango wants points; our widgets take pixels, so no conversion happens
    /// here — `font-size: 16px` is 16.
    pub font_size: Option<f32>,
    pub font_weight: Option<i32>,
    pub font_style: Option<FontStyle>,
    pub text_align: Option<TextAlign>,
    pub underline: Option<bool>,
    pub line_through: Option<bool>,
    pub line_height: Option<f32>,
    /// `cursor` — what the pointer looks like over this box or run.
    pub cursor: Option<Cursor>,
    /// `text-transform` — inherited, so a rule on a container re-cases the text
    /// of every descendant run.
    pub text_transform: Option<TextTransform>,
    /// Whether text wraps — INHERITED, like the rest of the text axis, so a
    /// container can set it once for everything inside it.
    pub white_space: Option<WhiteSpace>,
    /// `float` — out of flow, shifted to one side, line boxes shortened.
    /// NOT inherited: a float is a property of the box, not of its text.
    pub float: Option<Float>,
    /// `clear` — start below the floats on the named side.
    pub clear: Option<Clear>,
    /// `letter-spacing`, in **pixels**. CSS `normal` is `0` — the property adds
    /// to the font's own advance rather than replacing it, so zero means "the
    /// font decides", which is exactly what `normal` says.
    pub letter_spacing: Option<f32>,
}

impl CssProperties {
    /// Layer `other` on top of `self`: anything `other` specifies wins.
    pub fn merge(&mut self, other: &CssProperties) {
        macro_rules! take {
            ($($field:ident),+ $(,)?) => {$(
                if other.$field.is_some() { self.$field = other.$field.clone(); }
            )+};
        }
        take!(
            display,
            position,
            z_index,
            overflow,
            flex_direction,
            flex_wrap,
            justify_content,
            align_items,
            align_content,
            gap,
            row_gap,
            column_gap,
            border_collapse,
            border_spacing,
            table_layout,
            grid_template_columns,
            grid_template_rows,
            grid_template_areas,
            grid_auto_flow,
            grid_auto_rows,
            grid_auto_columns,
            justify_items,
            justify_self,
            grid_area,
            grid_column_start,
            grid_column_end,
            grid_row_start,
            grid_row_end,
            flex_grow,
            flex_shrink,
            flex_basis,
            align_self,
            order,
            width,
            height,
            min_width,
            min_height,
            max_width,
            max_height,
            border_radius,
            box_sizing,
            color,
            background_color,
            opacity,
            font_family,
            font_size,
            font_weight,
            font_style,
            text_align,
            underline,
            line_through,
            line_height,
            text_transform,
            white_space,
            cursor,
            letter_spacing,
        );
        self.offsets.merge_from(&other.offsets);
        self.margin.merge_from(&other.margin);
        self.padding.merge_from(&other.padding);
        self.border_width.merge_from(&other.border_width);
        self.border_style.merge_from(&other.border_style);
        self.border_color.merge_from(&other.border_color);
        if other.display_none {
            self.display_none = true;
        }
        if other.visibility_hidden {
            self.visibility_hidden = true;
        }
    }

    /// Is this box laid out by its container, or does it place itself?
    ///
    /// `position: absolute` and `fixed` are out of flow — the container hands
    /// them nothing and their own coordinates stand. This is the single question
    /// a container has to ask before arranging a child.
    pub fn is_out_of_flow(&self) -> bool {
        matches!(
            self.position,
            Some(Position::Absolute) | Some(Position::Fixed)
        )
    }

    /// Does this box arrange its children along an axis?
    pub fn is_flex_container(&self) -> bool {
        self.display == Some(Display::Flex)
    }

    /// The subset a child starts from — CSS inheritance.
    ///
    /// Everything else is dropped, which is the whole point: `width`,
    /// `position` and `background-color` are per-box answers and a child that
    /// picked them up from its parent would be nested inside a copy of it.
    ///
    /// **Not here, deliberately:**
    ///
    /// - `underline` / `line_through`. `text-decoration` is *not* an inherited
    ///   property. A decoration drawn by an ancestor is painted across its
    ///   in-flow descendants by the inline formatting context — a different
    ///   mechanism with a visible difference: an inherited underline would be
    ///   re-drawn at each descendant's own baseline and colour, a propagated
    ///   one is drawn once, by the ancestor, in the ancestor's colour. Without
    ///   inline runs there is nothing here to propagate *through*, so the
    ///   honest answer is that it does not reach descendants yet.
    /// - `visibility_hidden`. It *is* inherited, but the field is a sticky
    ///   `bool` — `merge` can only ever turn it on. `visibility: visible` on
    ///   the child of a hidden parent is the one thing the property is for,
    ///   and a one-way rule cannot express it. It needs an `Option` first.
    /// - `opacity`. Not inherited: descendants render *into* the ancestor's
    ///   opacity group rather than each taking the value.
    pub fn inherited(&self) -> CssProperties {
        CssProperties {
            color: self.color,
            font_family: self.font_family.clone(),
            font_size: self.font_size,
            font_weight: self.font_weight,
            font_style: self.font_style,
            text_align: self.text_align,
            line_height: self.line_height,
            text_transform: self.text_transform,
            white_space: self.white_space,
            cursor: self.cursor,
            letter_spacing: self.letter_spacing,
            border_collapse: self.border_collapse,
            border_spacing: self.border_spacing,
            ..CssProperties::default()
        }
    }

    /// The inherited subset written back out as declarations.
    ///
    /// The widget is told about a property by being handed a declaration, so an
    /// *inherited* value — which by definition was never declared on this
    /// element — has to be spelled to reach it. Serialising the computed value
    /// is what a browser's `getComputedStyle` does, and it keeps one channel to
    /// the widget instead of a second, typed one beside it.
    pub fn inherited_declarations(&self) -> Vec<(&'static str, String)> {
        let mut out = Vec::new();
        if let Some(color) = self.color {
            out.push(("color", serialize_color(color)));
        }
        if let Some(family) = &self.font_family {
            out.push(("font-family", family.clone()));
        }
        if let Some(size) = self.font_size {
            out.push(("font-size", format!("{size}px")));
        }
        if let Some(weight) = self.font_weight {
            out.push(("font-weight", weight.to_string()));
        }
        if let Some(style) = self.font_style {
            out.push((
                "font-style",
                match style {
                    FontStyle::Italic => "italic",
                    FontStyle::Oblique => "oblique",
                    FontStyle::Normal => "normal",
                }
                .to_string(),
            ));
        }
        if let Some(transform) = self.text_transform {
            out.push(("text-transform", transform.as_css().to_string()));
        }
        if let Some(cursor) = self.cursor {
            out.push(("cursor", cursor.as_css().to_string()));
        }
        if let Some(spacing) = self.letter_spacing {
            out.push(("letter-spacing", format!("{spacing}px")));
        }
        if let Some(align) = self.text_align {
            out.push((
                "text-align",
                match align {
                    TextAlign::Left => "left",
                    TextAlign::Center => "center",
                    TextAlign::Right => "right",
                    TextAlign::Justify => "justify",
                }
                .to_string(),
            ));
        }
        if let Some(height) = self.line_height {
            out.push(("line-height", format!("{height}px")));
        }
        if let Some(collapse) = self.border_collapse {
            out.push(("border-collapse", collapse.as_css().to_string()));
        }
        if let Some(spacing) = self.border_spacing {
            out.push(("border-spacing", spacing.to_string()));
        }
        out
    }
}

/// Every declaration in which two computed styles differ.
///
/// **This is what replaces asking which declarations might have changed.** The
/// old push re-applied a *guessed* set — the author declarations mentioning
/// `var(`, plus every inherited property — and did it whether or not anything
/// had actually moved. Both guesses exist only because there was no computed
/// value to compare against; with one, the question "what changed" has a real
/// answer and the guesses can go.
///
/// It is strictly more precise in both directions: a `--gap` rewritten to the
/// value it already had pushes nothing, and a `padding: var(--gap)` that really
/// did move is caught without `padding` having to be on any list.
///
/// The name/value pairs are what `apply_style_property` accepts, because that
/// is still the only channel to a widget — a declaration. Serialising a
/// computed value back into one is what `getComputedStyle` does.
///
/// A property that went from set to UNSET emits an EMPTY value rather than
/// being skipped. Losing a declaration — deleting the `--brand` that a
/// `color: var(--brand)` read, say — must take the element back to what it
/// would have been without it, not leave the widget frozen on the last thing it
/// was told. The box arms act on it because they read the resolved edge cache
/// rather than this string; the arms that parse the value instead (`color`,
/// `left`) return early on empty and do keep their last value, which is the
/// pre-existing "an empty declaration does not reset the widget" hole and is
/// not made worse here.
pub fn changed_declarations(old: &CssProperties, new: &CssProperties) -> Vec<(&'static str, String)> {
    let mut out: Vec<(&'static str, String)> = Vec::new();
    /// A field whose CSS spelling is its `Display` impl.
    macro_rules! changed {
        ($($name:literal => $field:ident),+ $(,)?) => {$(
            if old.$field != new.$field {
                out.push(($name, match &new.$field {
                    Some(v) => v.to_string(),
                    None => String::new(),
                }));
            }
        )+};
    }
    /// A field that needs a unit, a channel order, or a keyword table.
    macro_rules! changed_as {
        ($($name:literal => $field:ident, $how:expr);+ $(;)?) => {$(
            if old.$field != new.$field {
                out.push(($name, match &new.$field {
                    Some(v) => ($how)(v),
                    None => String::new(),
                }));
            }
        )+};
    }
    // Paint and text. These reach a widget with no layout consequence, which is
    // why they were the only ones the guessed push dared to re-apply.
    changed!("opacity" => opacity, "font-family" => font_family, "font-weight" => font_weight);
    changed_as!(
        "color" => color, |v: &u32| serialize_color(*v);
        "background-color" => background_color, |v: &u32| serialize_color(*v);
        "font-size" => font_size, |v: &f32| format!("{v}px");
        "font-style" => font_style, |v: &FontStyle| v.as_css().to_string();
        "text-align" => text_align, |v: &TextAlign| v.as_css().to_string();
        "line-height" => line_height, |v: &f32| format!("{v}px");
        "text-transform" => text_transform, |v: &TextTransform| v.as_css().to_string();
        "white-space" => white_space, |v: &WhiteSpace| v.as_css().to_string();
        "float" => float, |v: &Float| v.as_css().to_string();
        "clear" => clear, |v: &Clear| v.as_css().to_string();
        "cursor" => cursor, |v: &Cursor| v.as_css().to_string();
        "letter-spacing" => letter_spacing, |v: &f32| format!("{v}px");
    );
    // Layout. Included deliberately, and safe for the reason the diff exists:
    // a value that did not move is not pushed, so a font change no longer puts
    // the container through a relayout — which is what the old blanket
    // exclusion of geometry was protecting against, achieved by measuring
    // instead of by guessing.
    changed!(
        "z-index" => z_index,
        "width" => width,
        "height" => height,
        "min-width" => min_width,
        "min-height" => min_height,
        "max-width" => max_width,
        "max-height" => max_height,
        "gap" => gap,
        "row-gap" => row_gap,
        "column-gap" => column_gap,
        "flex-grow" => flex_grow,
        "order" => order,
    );
    changed_as!(
        "display" => display, |v: &Display| v.as_css().to_string();
        "position" => position, |v: &Position| v.as_css().to_string();
        "overflow" => overflow, |v: &Overflow| v.as_css().to_string();
        "flex-direction" => flex_direction, |v: &FlexDirection| v.as_css().to_string();
        "flex-wrap" => flex_wrap, |v: &FlexWrap| v.as_css().to_string();
        "justify-content" => justify_content, |v: &JustifyContent| v.as_css().to_string();
        "align-items" => align_items, |v: &AlignItems| v.as_css().to_string();
        "align-content" => align_content, |v: &JustifyContent| v.as_css().to_string();
        "align-self" => align_self, |v: &AlignItems| v.as_css().to_string();
        "border-collapse" => border_collapse, |v: &BorderCollapse| v.as_css().to_string();
        "border-spacing" => border_spacing, |v: &Length| v.to_string();
        "table-layout" => table_layout, |v: &TableLayout| v.as_css().to_string();
        "grid-template-columns" => grid_template_columns, |v: &TrackTemplate| track_template_css(v);
        "grid-template-rows" => grid_template_rows, |v: &TrackTemplate| track_template_css(v);
        "grid-template-areas" => grid_template_areas, |v: &Vec<Vec<String>>| area_template_css(v);
        "grid-auto-flow" => grid_auto_flow, |v: &GridAutoFlow| v.as_css().to_string();
        "grid-auto-rows" => grid_auto_rows, |v: &TrackSize| v.as_css();
        "grid-auto-columns" => grid_auto_columns, |v: &TrackSize| v.as_css();
        "justify-items" => justify_items, |v: &AlignItems| v.as_css().to_string();
        "justify-self" => justify_self, |v: &AlignItems| v.as_css().to_string();
        "grid-area" => grid_area, |v: &String| v.clone();
        "grid-column-start" => grid_column_start, |v: &GridLine| v.as_css();
        "grid-column-end" => grid_column_end, |v: &GridLine| v.as_css();
        "grid-row-start" => grid_row_start, |v: &GridLine| v.as_css();
        "grid-row-end" => grid_row_end, |v: &GridLine| v.as_css();
    );
    // The border, in its three axes. Not in the loop below because each axis
    // is a different type — width is a number, style a keyword, colour a
    // packed int — and each needs its own spelling on the way out.
    //
    // Missing entirely before: `box_edges` consumed border widths for LAYOUT,
    // so a border moved the content box and then painted nothing. A `<div>`
    // has no border by default and CAN have one, and this is the channel it
    // was lacking.
    for (name, before, after) in [
        ("border-top-width", old.border_width.top, new.border_width.top),
        ("border-right-width", old.border_width.right, new.border_width.right),
        ("border-bottom-width", old.border_width.bottom, new.border_width.bottom),
        ("border-left-width", old.border_width.left, new.border_width.left),
    ] {
        if before != after {
            out.push((name, after.map(|v| format!("{v}px")).unwrap_or_default()));
        }
    }
    for (name, before, after) in [
        ("border-top-style", old.border_style.top, new.border_style.top),
        ("border-right-style", old.border_style.right, new.border_style.right),
        ("border-bottom-style", old.border_style.bottom, new.border_style.bottom),
        ("border-left-style", old.border_style.left, new.border_style.left),
    ] {
        if before != after {
            out.push((name, after.map(|v| v.as_css().to_string()).unwrap_or_default()));
        }
    }
    for (name, before, after) in [
        ("border-top-color", old.border_color.top, new.border_color.top),
        ("border-right-color", old.border_color.right, new.border_color.right),
        ("border-bottom-color", old.border_color.bottom, new.border_color.bottom),
        ("border-left-color", old.border_color.left, new.border_color.left),
    ] {
        if before != after {
            out.push((name, after.map(serialize_color).unwrap_or_default()));
        }
    }
    // Per-side groups. Each side is its own declaration — `margin-left` rather
    // than a `margin` shorthand — so a change on one side does not restate the
    // other three, and the longhand arms already exist.
    for (names, old_side, new_side) in [
        (
            ["top", "right", "bottom", "left"],
            old.offsets,
            new.offsets,
        ),
        (
            ["margin-top", "margin-right", "margin-bottom", "margin-left"],
            old.margin,
            new.margin,
        ),
        (
            [
                "padding-top",
                "padding-right",
                "padding-bottom",
                "padding-left",
            ],
            old.padding,
            new.padding,
        ),
    ] {
        for (name, (before, after)) in names.into_iter().zip([
            (old_side.top, new_side.top),
            (old_side.right, new_side.right),
            (old_side.bottom, new_side.bottom),
            (old_side.left, new_side.left),
        ]) {
            if before != after {
                out.push((
                    name,
                    after.map(|v| v.to_string()).unwrap_or_default(),
                ));
            }
        }
    }
    // Two booleans, which have no "unset" to skip and so are stated whenever
    // they flip — including back to false, which IS the reset the `Option`
    // fields above cannot express.
    if old.display_none != new.display_none {
        out.push((
            "display",
            if new.display_none {
                "none".to_string()
            } else {
                new.display.map(|d| d.as_css()).unwrap_or("block").to_string()
            },
        ));
    }
    if old.visibility_hidden != new.visibility_hidden {
        out.push((
            "visibility",
            if new.visibility_hidden {
                "hidden"
            } else {
                "visible"
            }
            .to_string(),
        ));
    }
    out
}

/// The property names [`CssProperties::inherited`] carries down.
///
/// Spelled as names as well as fields because the write side is a declaration
/// (`set_style_property("color", …)`) and has to know, before parsing anything,
/// whether the write reaches descendants.
pub const INHERITED_PROPERTIES: &[&str] = &[
    "color",
    "font",
    "font-family",
    "font-size",
    "font-weight",
    "font-style",
    "text-align",
    "line-height",
    "text-transform",
    "cursor",
    "letter-spacing",
    // §17.6: the border model is a property of the TABLE, but it is declared on
    // the table and read by the cells, so the spec makes it inherited rather
    // than making every cell walk up to find its table. `border-spacing` goes
    // with it for the same reason.
    "border-collapse",
    "border-spacing",
];

/// A packed `0xAARRGGBB` written back as a CSS colour.
///
/// Opaque colours serialise as `#rrggbb` rather than `rgba(…, 1)` because that
/// is what round-trips through [`parse_color`] without a float in the middle.
pub fn serialize_color(argb: u32) -> String {
    let (a, r, g, b) = (argb >> 24 & 0xFF, argb >> 16 & 0xFF, argb >> 8 & 0xFF, argb & 0xFF);
    if a == 0xFF {
        format!("#{r:02x}{g:02x}{b:02x}")
    } else {
        format!("rgba({r}, {g}, {b}, {})", a as f32 / 255.0)
    }
}

// ── The box model ───────────────────────────────────────────────────────────

/// Four resolved per-side values, in pixels.
///
/// [`Sides`] is the *declaration* — each side optional, because "not specified"
/// has to survive merging. This is the *used value*: every side has a number,
/// because layout cannot subtract a `None`.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct Edges {
    pub top: f32,
    pub right: f32,
    pub bottom: f32,
    pub left: f32,
}

impl Edges {
    pub fn uniform(value: f32) -> Self {
        Self {
            top: value,
            right: value,
            bottom: value,
            left: value,
        }
    }

    /// Total across — what these edges cost a box's width.
    pub fn horizontal(self) -> f32 {
        self.left + self.right
    }

    /// Total down — what these edges cost a box's height.
    pub fn vertical(self) -> f32 {
        self.top + self.bottom
    }

    pub fn is_zero(self) -> bool {
        self.top == 0.0 && self.right == 0.0 && self.bottom == 0.0 && self.left == 0.0
    }
}

/// What separates a box's four rectangles.
///
/// CSS gives every element four nested rectangles — margin, border, padding and
/// content — and which one a question is about is not a detail. The containing
/// block for an absolutely positioned child is the **padding** box; a
/// background paints the **border** box; a container arranges its children
/// inside its **content** box. Modelling one rectangle per element makes those
/// three answers the same answer, which is right only while every edge is zero.
///
/// The widget's own rect is always the **border box**. That is a fact about the
/// rect, not about `box-sizing`: `box-sizing` decides what a *declared* `width`
/// measures, and the rect it ends up in is the border box either way.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct BoxEdges {
    pub margin: Edges,
    pub border: Edges,
    pub padding: Edges,
}

impl BoxEdges {
    /// Resolve declared sides into used values against a containing-block width.
    ///
    /// `basis` is the containing block's **width** for both axes, which looks
    /// like a bug and is the spec: a percentage margin or padding resolves
    /// against the containing block's width even on the vertical sides (CSS 2.1
    /// §8.3, §8.4). It is what makes `padding: 10%` square.
    ///
    /// An unspecified side is `0`, an `auto` margin is `0` here — centring is a
    /// layout decision, and layout is the only place that knows the leftover
    /// space to split.
    pub fn resolve(props: &CssProperties, basis: f32) -> Self {
        Self {
            margin: resolve_sides(&props.margin, basis),
            border: border_edges(&props.border_width),
            padding: resolve_sides(&props.padding, basis),
        }
    }
}

fn resolve_sides(sides: &Sides<Length>, basis: f32) -> Edges {
    let side = |v: Option<Length>| v.and_then(|l| l.resolve(basis)).unwrap_or(0.0);
    Edges {
        top: side(sides.top),
        right: side(sides.right),
        bottom: side(sides.bottom),
        left: side(sides.left),
    }
}

fn border_edges(sides: &Sides<f32>) -> Edges {
    Edges {
        top: sides.top.unwrap_or(0.0),
        right: sides.right.unwrap_or(0.0),
        bottom: sides.bottom.unwrap_or(0.0),
        left: sides.left.unwrap_or(0.0),
    }
}

// ── The store ───────────────────────────────────────────────────────────────

/// An element's inline declarations — `element.style`.
///
/// Records what was set, verbatim, so it reads back verbatim. Property names are
/// lower-cased (CSS property names are ASCII case-insensitive); values are kept
/// exactly as written, because a value's serialisation is observable and
/// round-tripping `#FFF` as `rgb(255,255,255)` is a behaviour change nobody
/// asked for.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Style {
    declarations: BTreeMap<String, String>,
}

impl Style {
    pub fn new() -> Self {
        Self::default()
    }

    /// The key a property name is stored under.
    ///
    /// Ordinary property names are ASCII case-insensitive and fold to
    /// lowercase. **Custom property names do not** — `--Brand` and `--brand`
    /// are two different properties (CSS Variables §2), because their names are
    /// author-chosen identifiers rather than a fixed vocabulary. Folding them
    /// made `var(--Brand)` silently read whatever `--brand` held.
    fn key(name: &str) -> String {
        let name = name.trim();
        if is_custom_property(name) {
            name.to_string()
        } else {
            name.to_ascii_lowercase()
        }
    }

    /// `style.setProperty(name, value)`. An empty value removes the
    /// declaration, as the CSSOM specifies.
    pub fn set(&mut self, name: &str, value: &str) {
        let name = Self::key(name);
        let value = value.trim();
        if value.is_empty() {
            self.declarations.remove(&name);
        } else {
            self.declarations.insert(name, value.to_string());
        }
    }

    /// `style.getPropertyValue(name)` — `""` when not set, per the CSSOM.
    pub fn get(&self, name: &str) -> &str {
        self.declarations
            .get(&Self::key(name))
            .map(String::as_str)
            .unwrap_or("")
    }

    pub fn remove(&mut self, name: &str) -> Option<String> {
        self.declarations.remove(&Self::key(name))
    }

    pub fn is_empty(&self) -> bool {
        self.declarations.is_empty()
    }

    pub fn iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.declarations
            .iter()
            .map(|(k, v)| (k.as_str(), v.as_str()))
    }

    /// Parse the whole declaration block into the typed view.
    pub fn properties(&self) -> CssProperties {
        self.properties_in(FontContext::default())
    }

    /// The typed view in a known font context — what the cascade uses.
    pub fn properties_in(&self, ctx: FontContext) -> CssProperties {
        self.properties_resolved_in(&|value| Some(value.to_string()), ctx)
    }

    /// As [`Style::properties`], but every value passes through `resolve`
    /// first — which is where `var()` substitution happens.
    ///
    /// The typed view is the second reader of the store, and the one the box
    /// model uses: without this, `padding: var(--gap)` would reach the widget
    /// resolved and reach `BoxEdges` as the unparseable literal `var(--gap)`,
    /// so the element's own padding and its container's idea of it would
    /// disagree.
    ///
    /// A value the resolver rejects is an invalid declaration — skipped, not
    /// applied as written.
    pub fn properties_resolved(&self, resolve: &dyn Fn(&str) -> Option<String>) -> CssProperties {
        self.properties_resolved_in(resolve, FontContext::default())
    }

    /// The typed view, with `em`/`rem` resolved against a real font context.
    pub fn properties_resolved_in(
        &self,
        resolve: &dyn Fn(&str) -> Option<String>,
        ctx: FontContext,
    ) -> CssProperties {
        let mut props = CssProperties::default();
        for (name, value) in self.iter() {
            // A custom property is a value, not a declaration about the box.
            if is_custom_property(name) {
                continue;
            }
            let Some(value) = resolve(value) else { continue };
            apply_declaration(&mut props, name, &value, ctx);
        }
        props
    }

    /// `style.cssText` — the declarations as a CSS string.
    pub fn css_text(&self) -> String {
        self.iter()
            .map(|(k, v)| format!("{k}: {v}"))
            .collect::<Vec<_>>()
            .join("; ")
    }
}

// ── Custom properties ───────────────────────────────────────────────────────

/// Is this a custom property — an author-defined `--name`?
///
/// <https://drafts.csswg.org/css-variables/#defining-variables>
pub fn is_custom_property(name: &str) -> bool {
    name.trim().starts_with("--")
}

/// Does this value need [`substitute_vars`] before it can be read as a value?
///
/// Worth asking separately: substitution allocates, and the overwhelming
/// majority of declarations contain no `var()` at all.
pub fn references_var(value: &str) -> bool {
    value.contains("var(")
}

/// How deep `var()` fallbacks may nest before the value is treated as invalid.
///
/// A guard rather than a limit anyone should reach: `var(--a, var(--b, …))` is
/// a real pattern, and `--a: var(--a)` is a cycle the spec makes invalid at
/// computed-value time. Without a bound, the cycle is a hang.
const MAX_VAR_DEPTH: usize = 16;

/// Replace every `var(--name[, fallback])` in `value` using `lookup`.
///
/// `lookup` answers what `--name` holds for the element being resolved,
/// INCLUDING inherited values — custom properties inherit, so the answer often
/// comes from an ancestor, and only the caller knows the tree.
///
/// A reference with no value and no fallback makes the declaration invalid at
/// computed-value time (CSS Variables §3), which is `None` here. That is not
/// the same as substituting an empty string: `width: var(--nope)` must leave
/// the width alone, not set it to nothing.
pub fn substitute_vars(value: &str, lookup: &dyn Fn(&str) -> Option<String>) -> Option<String> {
    substitute_vars_at(value, lookup, 0)
}

fn substitute_vars_at(
    value: &str,
    lookup: &dyn Fn(&str) -> Option<String>,
    depth: usize,
) -> Option<String> {
    if depth > MAX_VAR_DEPTH {
        return None;
    }
    let Some(start) = value.find("var(") else {
        return Some(value.to_string());
    };
    // Find the matching close paren, counting nesting so a fallback that is
    // itself a `var()` — or any other function — does not end the reference
    // early.
    let open = start + "var(".len();
    let mut depth_parens = 1usize;
    let mut end = None;
    for (i, c) in value[open..].char_indices() {
        match c {
            '(' => depth_parens += 1,
            ')' => {
                depth_parens -= 1;
                if depth_parens == 0 {
                    end = Some(open + i);
                    break;
                }
            }
            _ => {}
        }
    }
    // An unclosed `var(` is a parse error, and a parse error drops the whole
    // declaration rather than half of it.
    let end = end?;

    let inside = &value[open..end];
    let (name, fallback) = match inside.split_once(',') {
        Some((name, fallback)) => (name.trim(), Some(fallback.trim())),
        None => (inside.trim(), None),
    };
    if !is_custom_property(name) {
        return None;
    }

    let replacement = match lookup(name) {
        Some(found) => found,
        // The fallback is itself a value, so it may contain further references.
        None => substitute_vars_at(fallback?, lookup, depth + 1)?,
    };

    let substituted = format!("{}{}{}", &value[..start], replacement, &value[end + 1..]);
    // The replacement can itself contain `var()` — `--a: var(--b)` is legal —
    // so the result is resolved again rather than returned as written.
    substitute_vars_at(&substituted, lookup, depth + 1)
}

// ── Parsing ─────────────────────────────────────────────────────────────────

/// **What `em` and `rem` are relative to.**
///
/// A font-relative length has no meaning without this, and until inheritance
/// existed there was nothing to fill it with — so both units resolved against a
/// hardcoded 16px and `rem` was a synonym for `em`. With a computed style that
/// actually carries `font-size` down the tree, that is a live bug: the UA sheet
/// declares `h1 { font-size: 2em; margin: 0.67em 0 }`, and a heading inside a
/// 10px container was 32px tall regardless.
///
/// The two are NOT the same basis, and neither is a constant:
///
/// - **`em`** is the element's OWN computed `font-size` — except on
///   `font-size` itself, where the value being computed cannot be its own
///   basis, so it is the PARENT's. That exception is the whole subtlety, and it
///   is why the cascade resolves in two passes.
/// - **`rem`** is the ROOT element's computed `font-size`, always, whatever
///   this element or its ancestors declare. That is the entire point of the
///   unit.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FontContext {
    pub em: f32,
    pub rem: f32,
}

impl Default for FontContext {
    /// CSS's initial `font-size`, and the only honest answer with no element in
    /// hand.
    fn default() -> Self {
        FontContext {
            em: 16.0,
            rem: 16.0,
        }
    }
}

/// Parse a length. `%` and `auto` are preserved rather than resolved.
///
/// Font-relative units take the CSS initial 16px. Use [`parse_length_in`]
/// wherever an element's own font context is known.
pub fn parse_length(value: &str) -> Option<Length> {
    parse_length_in(value, FontContext::default())
}

/// Parse a length, resolving `em`/`rem` against a real font context.
pub fn parse_length_in(value: &str, ctx: FontContext) -> Option<Length> {
    let value = value.trim();
    if value.is_empty() {
        return None;
    }
    let lower = value.to_ascii_lowercase();
    match lower.as_str() {
        "auto" => return Some(Length::Auto),
        "inherit" | "initial" | "unset" | "none" => return None,
        _ => {}
    }

    let mut end = 0;
    for (i, ch) in value.char_indices() {
        if ch.is_ascii_digit() || ch == '.' || ((ch == '-' || ch == '+') && i == 0) {
            end = i + ch.len_utf8();
        } else {
            break;
        }
    }
    if end == 0 {
        return None;
    }
    let number: f32 = value[..end].parse().ok()?;
    match value[end..].trim().to_ascii_lowercase().as_str() {
        "%" => Some(Length::Percent(number)),
        "px" | "" => Some(Length::Px(number)),
        "pt" => Some(Length::Px(number * 96.0 / 72.0)),
        "em" => Some(Length::Px(number * ctx.em)),
        "rem" => Some(Length::Px(number * ctx.rem)),
        // An unknown unit is not a length. Guessing px here is how a typo
        // becomes a plausible-looking layout.
        _ => None,
    }
}

/// Pixels, when the value is an absolute length.
pub fn parse_px(value: &str) -> Option<f32> {
    parse_length(value).and_then(Length::px)
}

/// Expand a 1–4 value box shorthand into `[top, right, bottom, left]`.
pub fn expand_box_shorthand(value: &str) -> Sides<Length> {
    expand_box_shorthand_in(value, FontContext::default())
}

/// Expand a box shorthand in a known font context — `margin: 0.67em 0`.
pub fn expand_box_shorthand_in(value: &str, ctx: FontContext) -> Sides<Length> {
    let parts: Vec<Option<Length>> = value
        .split_whitespace()
        .map(|p| parse_length_in(p, ctx))
        .collect();
    let (top, right, bottom, left) = match parts.len() {
        1 => (parts[0], parts[0], parts[0], parts[0]),
        2 => (parts[0], parts[1], parts[0], parts[1]),
        3 => (parts[0], parts[1], parts[2], parts[1]),
        4 => (parts[0], parts[1], parts[2], parts[3]),
        _ => (None, None, None, None),
    };
    Sides {
        top,
        right,
        bottom,
        left,
    }
}

/// Parse `1px solid #000` in any order — width, style and colour are
/// distinguishable by shape, so CSS does not fix their order.
pub fn parse_border_shorthand(value: &str) -> (Option<f32>, Option<BorderStyle>, Option<u32>) {
    let mut width = None;
    let mut style = None;
    let mut color = None;
    for token in value.split_whitespace() {
        if let Some(s) = BorderStyle::parse(token) {
            style = Some(s);
        } else if let Some(px) = parse_px(token) {
            width = Some(px);
        } else if let Some(c) = parse_color(token) {
            color = Some(c);
        }
    }
    (width, style, color)
}

/// Parse the `font` shorthand: `[style] [weight] size[/line-height] family`.
pub fn parse_font_shorthand(value: &str, props: &mut CssProperties) {
    let mut rest = value.trim();
    loop {
        let token = rest.split_whitespace().next().unwrap_or("");
        if token.is_empty() {
            return;
        }
        let matched = if let Some(s) = FontStyle::parse(token) {
            props.font_style = Some(s);
            true
        } else if let Some(w) = parse_font_weight(token) {
            props.font_weight = Some(w);
            true
        } else if token.eq_ignore_ascii_case("small-caps") || token.eq_ignore_ascii_case("normal") {
            true
        } else {
            false
        };
        if !matched {
            break;
        }
        rest = rest[token.len()..].trim_start();
    }

    // size, optionally `size/line-height`
    let size_token = rest.split_whitespace().next().unwrap_or("");
    if size_token.is_empty() {
        return;
    }
    let (size, line_height) = match size_token.split_once('/') {
        Some((s, lh)) => (s, parse_px(lh)),
        None => (size_token, None),
    };
    if let Some(px) = parse_px(size) {
        props.font_size = Some(px);
    }
    if let Some(lh) = line_height {
        props.line_height = Some(lh);
    }

    let family = rest[size_token.len()..].trim();
    if !family.is_empty() {
        props.font_family = Some(first_font_family(family));
    }
}

fn parse_font_weight(token: &str) -> Option<i32> {
    match token.to_ascii_lowercase().as_str() {
        "bold" | "bolder" => Some(700),
        "lighter" => Some(300),
        other => match other.parse::<i32>() {
            Ok(w) if (1..=1000).contains(&w) => Some(w),
            _ => None,
        },
    }
}

fn first_font_family(value: &str) -> String {
    let first = value.split(',').next().unwrap_or(value).trim();
    first.trim_matches(['\'', '"']).to_string()
}

/// Parse a colour to packed `0xAARRGGBB`.
///
/// Handles `#rgb`, `#rgba`, `#rrggbb`, `#rrggbbaa`, `rgb()`/`rgba()`, and the
/// named colours the frontends actually emit. `transparent` is a fully
/// transparent black, as CSS defines it.
pub fn parse_color(value: &str) -> Option<u32> {
    let value = value.trim();
    if let Some(hex) = value.strip_prefix('#') {
        let expand = |c: u8| {
            let d = (c as char).to_digit(16)? as u32;
            Some(d * 17)
        };
        let bytes = hex.as_bytes();
        return match bytes.len() {
            3 | 4 => {
                let r = expand(bytes[0])?;
                let g = expand(bytes[1])?;
                let b = expand(bytes[2])?;
                let a = if bytes.len() == 4 {
                    expand(bytes[3])?
                } else {
                    255
                };
                Some(a << 24 | r << 16 | g << 8 | b)
            }
            6 | 8 => {
                let v = u32::from_str_radix(hex, 16).ok()?;
                if bytes.len() == 6 {
                    Some(0xFF00_0000 | v)
                } else {
                    Some(v.rotate_right(8))
                }
            }
            _ => None,
        };
    }

    let lower = value.to_ascii_lowercase();
    if let Some(args) = lower
        .strip_prefix("rgba(")
        .or_else(|| lower.strip_prefix("rgb("))
        .and_then(|a| a.strip_suffix(')'))
    {
        let parts: Vec<&str> = args
            .split([',', '/', ' '])
            .filter(|p| !p.is_empty())
            .collect();
        if parts.len() < 3 {
            return None;
        }
        let channel = |s: &str| -> Option<u32> {
            let s = s.trim();
            match s.strip_suffix('%') {
                Some(p) => p.parse::<f32>().ok().map(|v| (v * 2.55).round() as u32),
                None => s.parse::<f32>().ok().map(|v| v.round() as u32),
            }
            .map(|v| v.min(255))
        };
        let r = channel(parts[0])?;
        let g = channel(parts[1])?;
        let b = channel(parts[2])?;
        let a = match parts.get(3) {
            Some(a) => {
                let a = a.trim();
                match a.strip_suffix('%') {
                    Some(p) => (p.parse::<f32>().ok()? * 2.55).round() as u32,
                    None => (a.parse::<f32>().ok()? * 255.0).round() as u32,
                }
                .min(255)
            }
            None => 255,
        };
        return Some(a << 24 | r << 16 | g << 8 | b);
    }

    // A packed integer, which CSS has no syntax for and every toolkit uses:
    // Flutter's `Color.value`, WinForms' `Color.ToArgb`, VCL's `TColor`. The
    // channel order is **ARGB**, not the `#RRGGBBAA` above — that is not an
    // inconsistency to tidy away, it is what those APIs hand us, and reading
    // one as the other silently swaps a colour's alpha with its red.
    if let Some(hex) = lower.strip_prefix("0x") {
        return u32::from_str_radix(hex, 16).ok().map(opaque_if_no_alpha);
    }
    if let Ok(packed) = lower.parse::<u32>() {
        return Some(opaque_if_no_alpha(packed));
    }

    let rgb = match lower.as_str() {
        "transparent" => return Some(0),
        // Not CSS named colours, but spelled by the toolkits above and by
        // enough author markup to be worth answering.
        "lightgray" | "lightgrey" => 0xD3D3D3,
        "darkgray" | "darkgrey" => 0xA9A9A9,
        "black" => 0x000000,
        "silver" => 0xC0C0C0,
        "gray" | "grey" => 0x808080,
        "white" => 0xFFFFFF,
        "maroon" => 0x800000,
        "red" => 0xFF0000,
        "purple" => 0x800080,
        "fuchsia" | "magenta" => 0xFF00FF,
        "green" => 0x008000,
        "lime" => 0x00FF00,
        "olive" => 0x808000,
        "yellow" => 0xFFFF00,
        "navy" => 0x000080,
        "blue" => 0x0000FF,
        "teal" => 0x008080,
        "aqua" | "cyan" => 0x00FFFF,
        "orange" => 0xFFA500,
        _ => return None,
    };
    Some(0xFF00_0000 | rgb)
}

/// A packed value whose alpha byte is zero means opaque, not invisible.
///
/// `0x2196F3` is a colour, not "fully transparent black" — no caller writing a
/// six-digit constant means alpha 0.
pub(crate) fn opaque_if_no_alpha(packed: u32) -> u32 {
    if packed >> 24 == 0 {
        0xFF00_0000 | packed
    } else {
        packed
    }
}

/// Split a stylesheet into `(selector text, declaration block)` pairs.
///
/// The **syntax** half of a stylesheet, and deliberately only that: what a
/// selector means is `selector.rs`'s business and what a declaration means is
/// the rest of this file's, so this knows about braces, comments and at-rules
/// and nothing else.
///
/// Per CSS Syntax Level 3, an at-rule is skipped as a unit — `@media` and
/// `@supports` take a block, `@import` and `@charset` end at the semicolon.
/// Skipping is the compliant answer for a rule we do not implement: applying
/// the INSIDE of an unsupported `@media` would apply styles whose condition was
/// never evaluated, which is worse than ignoring it.
pub fn parse_rules(css: &str) -> Vec<(String, String)> {
    let text = strip_comments(css);
    let chars: Vec<char> = text.chars().collect();
    let mut rules = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        while i < chars.len() && chars[i].is_whitespace() {
            i += 1;
        }
        if i >= chars.len() {
            break;
        }
        if chars[i] == '@' {
            i = skip_at_rule(&chars, i);
            continue;
        }
        let prelude_start = i;
        while i < chars.len() && chars[i] != '{' && chars[i] != '}' {
            i += 1;
        }
        if i >= chars.len() || chars[i] == '}' {
            // A prelude with no block is not a rule. Dropping it is what the
            // spec's error handling says to do.
            i += 1;
            continue;
        }
        let selector: String = chars[prelude_start..i].iter().collect();
        i += 1; // past '{'
        let block_start = i;
        let mut depth = 1;
        while i < chars.len() && depth > 0 {
            match chars[i] {
                '{' => depth += 1,
                '}' => depth -= 1,
                _ => {}
            }
            if depth > 0 {
                i += 1;
            }
        }
        let block: String = chars[block_start..i.min(chars.len())].iter().collect();
        i += 1; // past '}'
        let selector = selector.trim().to_string();
        if !selector.is_empty() {
            rules.push((selector, block));
        }
    }
    rules
}

/// Skip a whole at-rule, block or semicolon terminated, and answer where it
/// ended.
fn skip_at_rule(chars: &[char], mut i: usize) -> usize {
    while i < chars.len() {
        match chars[i] {
            ';' => return i + 1,
            '{' => {
                let mut depth = 0;
                while i < chars.len() {
                    match chars[i] {
                        '{' => depth += 1,
                        '}' => {
                            depth -= 1;
                            if depth == 0 {
                                return i + 1;
                            }
                        }
                        _ => {}
                    }
                    i += 1;
                }
                return i;
            }
            _ => i += 1,
        }
    }
    i
}

/// Remove `/* … */`, which may appear anywhere a token may.
fn strip_comments(css: &str) -> String {
    let mut out = String::with_capacity(css.len());
    let bytes: Vec<char> = css.chars().collect();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == '/' && i + 1 < bytes.len() && bytes[i + 1] == '*' {
            i += 2;
            while i + 1 < bytes.len() && !(bytes[i] == '*' && bytes[i + 1] == '/') {
                i += 1;
            }
            i = (i + 2).min(bytes.len());
            // A comment is whitespace, not nothing: `a/**/b` is two tokens.
            out.push(' ');
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }
    out
}

/// Parse a declaration block (`a: b; c: d`) into the typed view.
pub fn parse_declarations(block: &str) -> CssProperties {
    let mut props = CssProperties::default();
    for declaration in block.split(';') {
        let Some((name, value)) = declaration.split_once(':') else {
            continue;
        };
        let name = name.trim().to_ascii_lowercase();
        let value = value.trim();
        if value.is_empty() {
            continue;
        }
        apply_declaration(&mut props, &name, value, FontContext::default());
    }
    props
}

/// Apply one declaration. `name` must already be lower-cased.
fn apply_declaration(props: &mut CssProperties, name: &str, value: &str, ctx: FontContext) {
    // Shadowed on purpose. Every length in this function is font-relative-aware
    // by virtue of being parsed here, rather than by thirty call sites each
    // having to remember to pass a context — and a new declaration added below
    // gets it for free instead of silently taking the 16px default.
    let parse_length = |v: &str| parse_length_in(v, ctx);
    let expand_box_shorthand = |v: &str| expand_box_shorthand_in(v, ctx);
    // `parse_px` needs the same treatment and is easy to miss precisely because
    // it looks like a unit conversion rather than a cascade question — it is
    // what `font-size` itself goes through, so leaving it unshadowed made the
    // one declaration that MUST see the parent's size the one that did not.
    let parse_px = |v: &str| parse_length_in(v, ctx).and_then(Length::px);
    match name {
        // ── Layout mode ──
        "display" => {
            if value.eq_ignore_ascii_case("none") {
                props.display_none = true;
            } else {
                props.display = Display::parse(value);
            }
        }
        "position" => props.position = Position::parse(value),
        "top" => props.offsets.top = parse_length(value),
        "right" => props.offsets.right = parse_length(value),
        "bottom" => props.offsets.bottom = parse_length(value),
        "left" => props.offsets.left = parse_length(value),
        "z-index" => props.z_index = value.trim().parse().ok(),
        "overflow" | "overflow-x" | "overflow-y" => props.overflow = Overflow::parse(value),

        // ── Flex container ──
        "flex-direction" => props.flex_direction = FlexDirection::parse(value),
        "flex-wrap" => props.flex_wrap = FlexWrap::parse(value),
        "flex-flow" => {
            for token in value.split_whitespace() {
                if let Some(d) = FlexDirection::parse(token) {
                    props.flex_direction = Some(d);
                } else if let Some(w) = FlexWrap::parse(token) {
                    props.flex_wrap = Some(w);
                }
            }
        }
        "justify-content" => props.justify_content = JustifyContent::parse(value),
        "align-items" => props.align_items = AlignItems::parse(value),
        // Shares `JustifyContent`'s vocabulary because the two distribute the
        // same way — one along the main axis, the other across the lines.
        "align-content" => props.align_content = JustifyContent::parse(value),
        "align-self" => props.align_self = AlignItems::parse(value),
        "gap" | "grid-gap" => {
            props.gap = parse_length(value);
            props.row_gap = props.gap;
            props.column_gap = props.gap;
        }
        "row-gap" => props.row_gap = parse_length(value),
        "column-gap" => props.column_gap = parse_length(value),
        "border-collapse" => props.border_collapse = BorderCollapse::parse(value),
        "border-spacing" => {
            // Two lengths are legal (`border-spacing: 4px 8px`); one field
            // holds one number, so the horizontal one wins and the vertical is
            // the same. Stated because a two-value table WILL look wrong rather
            // than fail, and that is the kind of divergence worth naming.
            props.border_spacing = parse_length(value.split_whitespace().next().unwrap_or(value));
        }
        "table-layout" => props.table_layout = TableLayout::parse(value),
        "grid-template-columns" => props.grid_template_columns = parse_track_template(value),
        "grid-template-rows" => props.grid_template_rows = parse_track_template(value),
        "grid-template-areas" => props.grid_template_areas = parse_area_template(value),
        "grid-auto-flow" => {
            // `dense` is a separate keyword that may precede or follow the
            // axis, so normalise before matching rather than listing four
            // spellings twice.
            let v = value.to_ascii_lowercase();
            let dense = v.split_whitespace().any(|w| w == "dense");
            let column = v.split_whitespace().any(|w| w == "column");
            props.grid_auto_flow = Some(match (column, dense) {
                (true, true) => GridAutoFlow::ColumnDense,
                (true, false) => GridAutoFlow::Column,
                (false, true) => GridAutoFlow::RowDense,
                (false, false) => GridAutoFlow::Row,
            });
        }
        "grid-auto-rows" => props.grid_auto_rows = TrackSize::parse(value),
        "grid-auto-columns" => props.grid_auto_columns = TrackSize::parse(value),
        "justify-items" => props.justify_items = AlignItems::parse(value),
        "justify-self" => props.justify_self = AlignItems::parse(value),
        // The `place-*` shorthands are `<align> <justify>`, in that order —
        // block axis first, which is the opposite of most CSS pairs and the
        // reason they are worth spelling out rather than guessing.
        "place-items" | "place-self" => {
            let mut parts = value.split_whitespace();
            let first = parts.next().unwrap_or("");
            let second = parts.next().unwrap_or(first);
            if name == "place-items" {
                props.align_items = AlignItems::parse(first);
                props.justify_items = AlignItems::parse(second);
            } else {
                props.align_self = AlignItems::parse(first);
                props.justify_self = AlignItems::parse(second);
            }
        }
        "place-content" => {
            let mut parts = value.split_whitespace();
            let first = parts.next().unwrap_or("");
            let second = parts.next().unwrap_or(first);
            props.align_content = JustifyContent::parse(first);
            props.justify_content = JustifyContent::parse(second);
        }
        // `grid-area` is two properties wearing one name. A single identifier
        // is an AREA; anything with a `/` is the four-line shorthand, and the
        // two have nothing in common but their spelling.
        "grid-area" => {
            if value.contains('/') {
                let parts: Vec<&str> = value.split('/').collect();
                props.grid_row_start = parts.first().and_then(|v| GridLine::parse(v));
                props.grid_column_start = parts.get(1).and_then(|v| GridLine::parse(v));
                props.grid_row_end = parts.get(2).and_then(|v| GridLine::parse(v));
                props.grid_column_end = parts.get(3).and_then(|v| GridLine::parse(v));
            } else {
                let name = value.trim();
                props.grid_area = (!name.is_empty()).then(|| name.to_string());
            }
        }
        "grid-column" => {
            let (start, end) = parse_grid_placement(value);
            props.grid_column_start = start;
            props.grid_column_end = end;
        }
        "grid-row" => {
            let (start, end) = parse_grid_placement(value);
            props.grid_row_start = start;
            props.grid_row_end = end;
        }
        "grid-column-start" => props.grid_column_start = GridLine::parse(value),
        "grid-column-end" => props.grid_column_end = GridLine::parse(value),
        "grid-row-start" => props.grid_row_start = GridLine::parse(value),
        "grid-row-end" => props.grid_row_end = GridLine::parse(value),

        // ── Flex item ──
        // `flex: <grow> [shrink] [basis]`, plus the keyword forms. `flex: 1` is
        // `1 1 0%` — the basis matters: it is the difference between "share the
        // space" and "share what is left after content".
        "flex" => match value.trim().to_ascii_lowercase().as_str() {
            "none" => {
                props.flex_grow = Some(0.0);
                props.flex_shrink = Some(0.0);
                props.flex_basis = Some(Length::Auto);
            }
            "auto" => {
                props.flex_grow = Some(1.0);
                props.flex_shrink = Some(1.0);
                props.flex_basis = Some(Length::Auto);
            }
            "initial" => {
                props.flex_grow = Some(0.0);
                props.flex_shrink = Some(1.0);
                props.flex_basis = Some(Length::Auto);
            }
            _ => {
                let tokens: Vec<&str> = value.split_whitespace().collect();
                match tokens.len() {
                    1 => {
                        if let Ok(grow) = tokens[0].parse::<f32>() {
                            props.flex_grow = Some(grow);
                            props.flex_shrink = Some(1.0);
                            props.flex_basis = Some(Length::Percent(0.0));
                        } else if let Some(basis) = parse_length(tokens[0]) {
                            props.flex_grow = Some(1.0);
                            props.flex_shrink = Some(1.0);
                            props.flex_basis = Some(basis);
                        }
                    }
                    2 => {
                        props.flex_grow = tokens[0].parse().ok();
                        match tokens[1].parse::<f32>() {
                            Ok(shrink) => props.flex_shrink = Some(shrink),
                            Err(_) => props.flex_basis = parse_length(tokens[1]),
                        }
                    }
                    _ => {
                        props.flex_grow = tokens[0].parse().ok();
                        props.flex_shrink = tokens[1].parse().ok();
                        props.flex_basis = parse_length(tokens[2]);
                    }
                }
            }
        },
        "flex-grow" => props.flex_grow = value.trim().parse().ok(),
        "flex-shrink" => props.flex_shrink = value.trim().parse().ok(),
        "flex-basis" => props.flex_basis = parse_length(value),
        "order" => props.order = value.trim().parse().ok(),

        // ── Box ──
        "width" => props.width = parse_length(value),
        "height" => props.height = parse_length(value),
        "min-width" => props.min_width = parse_length(value),
        "min-height" => props.min_height = parse_length(value),
        "max-width" => props.max_width = parse_length(value),
        "max-height" => props.max_height = parse_length(value),
        "margin" => props.margin.merge_from(&expand_box_shorthand(value)),
        "padding" => props.padding.merge_from(&expand_box_shorthand(value)),
        "margin-top" | "margin-block-start" => props.margin.top = parse_length(value),
        "margin-right" | "margin-inline-end" => props.margin.right = parse_length(value),
        "margin-bottom" | "margin-block-end" => props.margin.bottom = parse_length(value),
        "margin-left" | "margin-inline-start" => props.margin.left = parse_length(value),
        "padding-top" | "padding-block-start" => props.padding.top = parse_length(value),
        "padding-right" | "padding-inline-end" => props.padding.right = parse_length(value),
        "padding-bottom" | "padding-block-end" => props.padding.bottom = parse_length(value),
        "padding-left" | "padding-inline-start" => props.padding.left = parse_length(value),

        // ── Border ──
        "border" => {
            let (w, s, c) = parse_border_shorthand(value);
            if let Some(w) = w {
                props.border_width.set_all(w);
            }
            if let Some(s) = s {
                props.border_style.set_all(s);
            }
            if let Some(c) = c {
                props.border_color.set_all(c);
            }
        }
        "border-top" | "border-right" | "border-bottom" | "border-left" => {
            let (w, s, c) = parse_border_shorthand(value);
            let side = name.rsplit('-').next().unwrap_or("");
            set_side(&mut props.border_width, side, w);
            set_side(&mut props.border_style, side, s);
            set_side(&mut props.border_color, side, c);
        }
        "border-width" => {
            let sides = expand_box_shorthand(value);
            props.border_width.merge_from(&Sides {
                top: sides.top.and_then(Length::px),
                right: sides.right.and_then(Length::px),
                bottom: sides.bottom.and_then(Length::px),
                left: sides.left.and_then(Length::px),
            });
        }
        "border-style" => {
            if let Some(s) = BorderStyle::parse(value) {
                props.border_style.set_all(s);
            }
        }
        "border-color" => {
            if let Some(c) = parse_color(value) {
                props.border_color.set_all(c);
            }
        }
        "border-radius" => props.border_radius = parse_px(value),
        "box-sizing" => props.box_sizing = BoxSizing::parse(value),

        // ── Paint ──
        "color" => props.color = parse_color(value),
        "background-color" | "background" => props.background_color = parse_color(value),
        "opacity" => props.opacity = value.trim().parse().ok(),
        "visibility" => props.visibility_hidden = value.eq_ignore_ascii_case("hidden"),

        // ── Text ──
        "font-family" => props.font_family = Some(first_font_family(value)),
        "font-size" => props.font_size = parse_px(value),
        "font-weight" => props.font_weight = parse_font_weight(value).or(Some(400)),
        "font-style" => props.font_style = FontStyle::parse(value),
        "font" => parse_font_shorthand(value, props),
        "text-align" => {
            props.text_align = match value.trim().to_ascii_lowercase().as_str() {
                "start" => Some(TextAlign::Left),
                "end" => Some(TextAlign::Right),
                other => TextAlign::parse(other),
            }
        }
        "text-decoration" | "text-decoration-line" => {
            let lower = value.to_ascii_lowercase();
            if lower.contains("none") {
                props.underline = Some(false);
                props.line_through = Some(false);
            } else {
                if lower.contains("underline") {
                    props.underline = Some(true);
                }
                if lower.contains("line-through") {
                    props.line_through = Some(true);
                }
            }
        }
        "text-transform" => props.text_transform = TextTransform::parse(value),
        "white-space" => props.white_space = WhiteSpace::parse(value),
        "float" => props.float = Float::parse(value),
        "clear" => props.clear = Clear::parse(value),
        "cursor" => props.cursor = Cursor::parse(value),
        // `normal` is the initial value and means zero ADDED spacing.
        "letter-spacing" => {
            props.letter_spacing = if value.trim().eq_ignore_ascii_case("normal") {
                Some(0.0)
            } else {
                parse_px(value)
            }
        }
        "line-height" => {
            let trimmed = value.trim();
            props.line_height = match trimmed.parse::<f32>() {
                // A unitless line-height is a multiple of the font size.
                Ok(multiple) => props.font_size.map(|s| s * multiple),
                Err(_) => parse_px(trimmed),
            };
        }

        // Anything else is still STORED by `Style`; it simply has no typed
        // meaning here. That is a rendering gap, not data loss.
        _ => {}
    }
}

fn set_side<T: Copy>(sides: &mut Sides<T>, side: &str, value: Option<T>) {
    let Some(value) = value else { return };
    match side {
        "top" => sides.top = Some(value),
        "right" => sides.right = Some(value),
        "bottom" => sides.bottom = Some(value),
        "left" => sides.left = Some(value),
        _ => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn store_reads_back_what_was_written() {
        // The defect this module exists to fix: a style write used to become a
        // widget command and the CSS was forgotten, so the read answered "".
        let mut style = Style::new();
        style.set("color", "red");
        style.set("Background-Color", "#fff");
        assert_eq!(style.get("color"), "red");
        assert_eq!(style.get("background-color"), "#fff");
        assert_eq!(style.get("padding"), "");
    }

    #[test]
    fn empty_value_removes_the_declaration() {
        let mut style = Style::new();
        style.set("color", "red");
        style.set("color", "");
        assert_eq!(style.get("color"), "");
        assert!(style.is_empty());
    }

    #[test]
    fn unknown_properties_survive_even_though_nothing_renders_them() {
        let mut style = Style::new();
        style.set("mix-blend-mode", "multiply");
        assert_eq!(style.get("mix-blend-mode"), "multiply");
        assert_eq!(style.properties(), CssProperties::default());
    }

    #[test]
    fn percentages_stay_symbolic() {
        // Resolving `%` at parse time is what the source did (`v * 0.16`);
        // it needs a containing block, which parsing does not have.
        assert_eq!(parse_length("50%"), Some(Length::Percent(50.0)));
        assert_eq!(parse_length("50%").and_then(Length::px), None);
        assert_eq!(Length::Percent(50.0).resolve(300.0), Some(150.0));
        assert_eq!(parse_length("8px"), Some(Length::Px(8.0)));
        assert_eq!(parse_length("auto"), Some(Length::Auto));
    }

    #[test]
    fn an_unknown_unit_is_not_a_length() {
        assert_eq!(parse_length("10furlongs"), None);
    }

    #[test]
    fn display_none_is_visibility_not_a_layout_mode() {
        // These were one property once: `display: flex` marked an element
        // visible and selected no layout.
        let hidden = parse_declarations("display: none");
        assert!(hidden.display_none);
        assert_eq!(hidden.display, None);

        let flex = parse_declarations("display: flex");
        assert!(!flex.display_none);
        assert_eq!(flex.display, Some(Display::Flex));
        assert!(flex.is_flex_container());
    }

    #[test]
    fn absolute_is_out_of_flow_and_static_is_not() {
        // What every pixel-positioned frontend means by setting Left/Top.
        assert!(parse_declarations("position: absolute").is_out_of_flow());
        assert!(parse_declarations("position: fixed").is_out_of_flow());
        assert!(!parse_declarations("position: static").is_out_of_flow());
        assert!(!CssProperties::default().is_out_of_flow());
    }

    #[test]
    fn box_shorthand_follows_the_one_to_four_value_rule() {
        let one = expand_box_shorthand("4px");
        assert_eq!(one.top, Some(Length::Px(4.0)));
        assert_eq!(one.left, Some(Length::Px(4.0)));

        let two = expand_box_shorthand("1px 2px");
        assert_eq!(two.top, Some(Length::Px(1.0)));
        assert_eq!(two.right, Some(Length::Px(2.0)));
        assert_eq!(two.bottom, Some(Length::Px(1.0)));
        assert_eq!(two.left, Some(Length::Px(2.0)));

        let three = expand_box_shorthand("1px 2px 3px");
        assert_eq!(three.bottom, Some(Length::Px(3.0)));
        assert_eq!(three.left, Some(Length::Px(2.0)));

        let four = expand_box_shorthand("1px 2px 3px 4px");
        assert_eq!(four.left, Some(Length::Px(4.0)));
    }

    #[test]
    fn border_shorthand_is_order_independent() {
        let (w, s, c) = parse_border_shorthand("1px solid #000");
        assert_eq!(w, Some(1.0));
        assert_eq!(s, Some(BorderStyle::Solid));
        assert_eq!(c, Some(0xFF00_0000));

        let (w, s, c) = parse_border_shorthand("dashed red 2px");
        assert_eq!(w, Some(2.0));
        assert_eq!(s, Some(BorderStyle::Dashed));
        assert_eq!(c, Some(0xFFFF_0000));
    }

    #[test]
    fn flex_one_is_grow_one_shrink_one_basis_zero() {
        let props = parse_declarations("flex: 1");
        assert_eq!(props.flex_grow, Some(1.0));
        assert_eq!(props.flex_shrink, Some(1.0));
        assert_eq!(props.flex_basis, Some(Length::Percent(0.0)));

        // A docked bar: fixed, does not grow, sized by content.
        let bar = parse_declarations("flex: 0 0 auto");
        assert_eq!(bar.flex_grow, Some(0.0));
        assert_eq!(bar.flex_shrink, Some(0.0));
        assert_eq!(bar.flex_basis, Some(Length::Auto));
    }

    #[test]
    fn colors_parse_to_packed_argb() {
        assert_eq!(parse_color("#fff"), Some(0xFFFF_FFFF));
        assert_eq!(parse_color("#ff0000"), Some(0xFFFF_0000));
        assert_eq!(parse_color("rgb(255, 0, 0)"), Some(0xFFFF_0000));
        assert_eq!(parse_color("rgba(255, 0, 0, 0.5)"), Some(0x80FF_0000));
        assert_eq!(parse_color("transparent"), Some(0));
        assert_eq!(parse_color("nonsense"), None);
    }

    #[test]
    fn font_shorthand_reads_style_weight_size_and_family() {
        let props = parse_declarations("font: italic bold 16px/24px 'Helvetica', sans-serif");
        assert_eq!(props.font_style, Some(FontStyle::Italic));
        assert_eq!(props.font_weight, Some(700));
        assert_eq!(props.font_size, Some(16.0));
        assert_eq!(props.line_height, Some(24.0));
        assert_eq!(props.font_family.as_deref(), Some("Helvetica"));
    }

    #[test]
    fn font_size_stays_in_pixels() {
        // The source converted to points for Pango; our widgets take pixels.
        assert_eq!(parse_declarations("font-size: 16px").font_size, Some(16.0));
    }

    #[test]
    fn merge_layers_later_declarations_over_earlier() {
        let mut base = parse_declarations("color: red; padding: 4px; display: flex");
        let over = parse_declarations("color: blue; padding-left: 12px");
        base.merge(&over);
        assert_eq!(base.color, parse_color("blue"));
        assert_eq!(base.padding.left, Some(Length::Px(12.0)));
        assert_eq!(base.padding.top, Some(Length::Px(4.0)));
        assert_eq!(base.display, Some(Display::Flex));
    }

    #[test]
    fn style_parses_its_own_declarations() {
        let mut style = Style::new();
        style.set("display", "flex");
        style.set("flex-direction", "column");
        style.set("justify-content", "space-between");
        let props = style.properties();
        assert!(props.is_flex_container());
        assert_eq!(props.flex_direction, Some(FlexDirection::Column));
        assert_eq!(props.justify_content, Some(JustifyContent::SpaceBetween));
    }
}
