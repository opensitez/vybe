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
    Display {
        Block => "block",
        Flex => "flex",
        InlineBlock => "inline-block",
        Inline => "inline",
        Grid => "grid",
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
    AlignItems {
        FlexStart => "flex-start",
        FlexEnd => "flex-end",
        Center => "center",
        Baseline => "baseline",
        Stretch => "stretch",
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
    pub gap: Option<Length>,

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
            gap,
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

    /// `style.setProperty(name, value)`. An empty value removes the
    /// declaration, as the CSSOM specifies.
    pub fn set(&mut self, name: &str, value: &str) {
        let name = name.trim().to_ascii_lowercase();
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
            .get(&name.trim().to_ascii_lowercase())
            .map(String::as_str)
            .unwrap_or("")
    }

    pub fn remove(&mut self, name: &str) -> Option<String> {
        self.declarations.remove(&name.trim().to_ascii_lowercase())
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
        let mut props = CssProperties::default();
        for (name, value) in self.iter() {
            apply_declaration(&mut props, name, value);
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

// ── Parsing ─────────────────────────────────────────────────────────────────

/// Parse a length. `%` and `auto` are preserved rather than resolved.
///
/// `em`/`rem` assume a 16px root, which is the CSS initial value and the only
/// answer available without a font context.
pub fn parse_length(value: &str) -> Option<Length> {
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
        "em" | "rem" => Some(Length::Px(number * 16.0)),
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
    let parts: Vec<Option<Length>> = value.split_whitespace().map(parse_length).collect();
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
                let a = if bytes.len() == 4 { expand(bytes[3])? } else { 255 };
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
        let parts: Vec<&str> = args.split([',', '/', ' ']).filter(|p| !p.is_empty()).collect();
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

    let rgb = match lower.as_str() {
        "transparent" => return Some(0),
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
        apply_declaration(&mut props, &name, value);
    }
    props
}

/// Apply one declaration. `name` must already be lower-cased.
fn apply_declaration(props: &mut CssProperties, name: &str, value: &str) {
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
        "align-self" => props.align_self = AlignItems::parse(value),
        "gap" | "grid-gap" | "row-gap" | "column-gap" => props.gap = parse_length(value),

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
