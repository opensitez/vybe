//! Flutter enum surface.
//!
//! Flutter enum constants are compile-time known, so the Dart frontend folds
//! `Clip.antiAlias` to its canonical `"Clip.antiAlias"` spelling — which is
//! exactly what Dart's `toString()` yields for an enum value, so identity
//! comparison (`a.clipBehavior == Clip.antiAlias`), printing and `switch` all
//! behave. `.name` and `.index` fold to the value's short name and ordinal.
//!
//! The DATA lives here with the rest of the catalog; the walker only consumes
//! it (`vybe_platform_flutter::emitter::flutter_enums`).

/// `(EnumName, values in declaration order)`. Ordinal = index in the slice, so
/// value order is significant — it is `.index`.
pub(crate) const ENUMS: &[(&str, &[&str])] = &[
    // ── Layout / flex ───────────────────────────────────────────────────
    ("Axis", &["horizontal", "vertical"]),
    (
        "AxisDirection",
        &["up", "right", "down", "left"],
    ),
    (
        "MainAxisAlignment",
        &["start", "end", "center", "spaceBetween", "spaceAround", "spaceEvenly"],
    ),
    ("MainAxisSize", &["min", "max"]),
    (
        "CrossAxisAlignment",
        &["start", "end", "center", "stretch", "baseline"],
    ),
    ("FlexFit", &["tight", "loose"]),
    ("StackFit", &["loose", "expand", "passthrough"]),
    ("VerticalDirection", &["up", "down"]),
    ("TextDirection", &["rtl", "ltr"]),
    ("GrowthDirection", &["forward", "reverse"]),
    ("ScrollDirection", &["idle", "forward", "reverse"]),
    // ── Painting / clipping ─────────────────────────────────────────────
    (
        "Clip",
        &["none", "hardEdge", "antiAlias", "antiAliasWithSaveLayer"],
    ),
    (
        "BoxFit",
        &["fill", "contain", "cover", "fitWidth", "fitHeight", "none", "scaleDown"],
    ),
    ("BoxShape", &["rectangle", "circle"]),
    ("DecorationPosition", &["background", "foreground"]),
    (
        "ImageRepeat",
        &["repeat", "repeatX", "repeatY", "noRepeat"],
    ),
    (
        "FilterQuality",
        &["none", "low", "medium", "high"],
    ),
    ("PaintingStyle", &["fill", "stroke"]),
    ("StrokeCap", &["butt", "round", "square"]),
    ("StrokeJoin", &["miter", "round", "bevel"]),
    ("PathFillType", &["nonZero", "evenOdd"]),
    (
        "PathOperation",
        &["difference", "intersect", "union", "xor", "reverseDifference"],
    ),
    ("BlurStyle", &["normal", "solid", "outer", "inner"]),
    ("TileMode", &["clamp", "repeated", "mirror", "decal"]),
    (
        "BlendMode",
        &[
            "clear", "src", "dst", "srcOver", "dstOver", "srcIn", "dstIn", "srcOut",
            "dstOut", "srcATop", "dstATop", "xor", "plus", "modulate", "screen",
            "overlay", "darken", "lighten", "colorDodge", "colorBurn", "hardLight",
            "softLight", "difference", "exclusion", "multiply", "hue", "saturation",
            "color", "luminosity",
        ],
    ),
    ("PixelFormat", &["rgba8888", "bgra8888"]),
    // ── Text ────────────────────────────────────────────────────────────
    (
        "TextAlign",
        &["left", "right", "center", "justify", "start", "end"],
    ),
    ("TextBaseline", &["alphabetic", "ideographic"]),
    (
        "TextOverflow",
        &["clip", "fade", "ellipsis", "visible"],
    ),
    (
        "TextDecoration",
        &["none", "underline", "overline", "lineThrough"],
    ),
    (
        "TextDecorationStyle",
        &["solid", "double", "dotted", "dashed", "wavy"],
    ),
    ("FontStyle", &["normal", "italic"]),
    (
        "FontWeight",
        &["w100", "w200", "w300", "w400", "w500", "w600", "w700", "w800", "w900"],
    ),
    ("TextWidthBasis", &["parent", "longestLine"]),
    // ── Material / widgets ──────────────────────────────────────────────
    ("StepState", &["indexed", "editing", "complete", "disabled", "error"]),
    ("StepperType", &["vertical", "horizontal"]),
    ("BottomNavigationBarType", &["fixed", "shifting"]),
    ("AutovalidateMode", &["disabled", "always", "onUserInteraction"]),
    ("HeroFlightDirection", &["push", "pop"]),
    (
        "DismissDirection",
        &[
            "vertical", "horizontal", "endToStart", "startToEnd", "up", "down", "none",
        ],
    ),
    ("DragStartBehavior", &["down", "start"]),
    (
        "HitTestBehavior",
        &["deferToChild", "opaque", "translucent"],
    ),
    ("RoutePopDisposition", &["pop", "doNotPop", "bubble"]),
    (
        "ConnectionState",
        &["none", "waiting", "active", "done"],
    ),
    ("Brightness", &["dark", "light"]),
    (
        "SemanticsAction",
        &[
            "tap", "longPress", "scrollLeft", "scrollRight", "scrollUp", "scrollDown",
            "increase", "decrease", "showOnScreen", "dismiss",
        ],
    ),
];

/// Every Flutter enum: `(EnumName, values in ordinal order)`.
pub fn flutter_enums() -> &'static [(&'static str, &'static [&'static str])] {
    ENUMS
}

/// The ordinal of `value` within `enum_name`, or `None` when either is unknown.
pub fn enum_value_index(enum_name: &str, value: &str) -> Option<usize> {
    ENUMS
        .iter()
        .find(|(n, _)| *n == enum_name)
        .and_then(|(_, vs)| vs.iter().position(|v| *v == value))
}
