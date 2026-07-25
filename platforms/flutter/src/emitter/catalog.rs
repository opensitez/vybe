//! Flutter widget catalog — shared adapter types + aggregation.
//!
//! Flutter-shaped adapter surface. Each entry lowers a Flutter widget onto
//! an existing `vybe_widgets`/`vybe:gui` control (the same host constructors
//! the dotnet WinForms and plib VCL adapters use) — no Flutter-specific host
//! functions. The catalog is pure DATA: resolution logic lives in the shared
//! namespace resolver; construction/field-capture/`is`-checks ride the shared
//! class machinery (`vybe_emitter::classes`).
//!
//! The actual widget entries live in per-category adapter modules under
//! [`crate::emitter::widgets`] (`layout`, `material`, `inputs`, `scrolling`,
//! `painting`, `gestures`, `builders`, `cupertino`, `value_types`, `keys`,
//! `images`, `animation`, `focus`, `abstracts`). This module owns only the
//! shared adapter TYPES, the constructor helpers those modules use, the
//! double-field/value-type type seed, and the aggregation that concatenates
//! every category into one lookup slice.

use std::sync::OnceLock;

/// A Flutter widget/type adapter entry.
#[derive(Debug, Clone, Copy)]
pub struct FlutterClass {
    pub name: &'static str,
    /// Single-inheritance parent (`is`/type-identity chain).
    pub parent: Option<&'static str>,
    /// Extra identity names for `is` checks beyond the parent chain
    /// (Flutter mixins/interfaces such as `Diagnosticable`).
    pub interfaces: &'static [&'static str],
    /// Named/positional constructor params captured as instance fields
    /// (`Scaffold(appBar:).appBar`, `Text('x').data`). Each becomes a
    /// readable field on the constructed object.
    pub fields: &'static [FlutterField],
    /// Backing `vybe:gui` control constructor (`new_Label`, `new_Panel`,
    /// `new_TreeView`, …). `None` for abstract bases and pure-data types.
    pub widget_host_fn: Option<&'static str>,
    /// A TRANSPARENT wrapper: the widget contributes no control of its own,
    /// only an effect on its child (`Opacity`, `ClipRRect`, `RepaintBoundary`,
    /// `Tooltip`, …). The realizer renders the child directly in the wrapper's
    /// place instead of creating a Panel that just nests one child — a bare
    /// Panel per wrapper is pure layout noise, since none of the effects
    /// (opacity/clip/transform) are expressible on the backing controls.
    pub transparent: bool,
}

impl FlutterClass {
    /// A concrete widget backed by a `vybe:gui` control constructor.
    pub(crate) const fn widget(
        name: &'static str,
        parent: &'static str,
        host: &'static str,
        fields: &'static [FlutterField],
    ) -> Self {
        FlutterClass {
            name,
            parent: Some(parent),
            interfaces: NO_INTERFACES,
            fields,
            widget_host_fn: Some(host),
            transparent: false,
        }
    }

    /// A TRANSPARENT wrapper widget: keeps full type identity and field
    /// capture, but realizes its child in its own place rather than creating
    /// a control. See [`FlutterClass::transparent`].
    pub(crate) const fn wrapper(
        name: &'static str,
        parent: &'static str,
        fields: &'static [FlutterField],
    ) -> Self {
        FlutterClass {
            name,
            parent: Some(parent),
            interfaces: NO_INTERFACES,
            fields,
            widget_host_fn: Some("Panel"),
            transparent: true,
        }
    }

    /// An abstract base in the `is`/identity chain (no backing control).
    pub(crate) const fn abstract_(name: &'static str, parent: Option<&'static str>) -> Self {
        FlutterClass {
            name,
            parent,
            interfaces: NO_INTERFACES,
            fields: NO_FIELDS,
            widget_host_fn: None,
            transparent: false,
        }
    }

    /// A pure data / value type (no backing control): Color, Offset, FocusNode…
    pub(crate) const fn data(
        name: &'static str,
        parent: Option<&'static str>,
        fields: &'static [FlutterField],
    ) -> Self {
        FlutterClass {
            name,
            parent,
            interfaces: NO_INTERFACES,
            fields,
            widget_host_fn: None,
            transparent: false,
        }
    }

    /// A data type carrying extra `is`-identity names (e.g. `Key` → ValueKey).
    pub(crate) const fn data_with_interfaces(
        name: &'static str,
        parent: Option<&'static str>,
        interfaces: &'static [&'static str],
        fields: &'static [FlutterField],
    ) -> Self {
        FlutterClass {
            name,
            parent,
            interfaces,
            fields,
            widget_host_fn: None,
            transparent: false,
        }
    }
}

/// A constructor param → instance field.
#[derive(Debug, Clone, Copy)]
pub struct FlutterField {
    pub name: &'static str,
    /// Positional constructor slot (`Text('hello')` → `data` at slot 0),
    /// or `None` for named-only params (`Scaffold(appBar: …)`).
    pub positional: Option<u8>,
    /// Default value expression source when the param is omitted
    /// (`Column().direction == Axis.vertical`). `None` → defaults to null.
    pub default: Option<&'static str>,
    /// True when the value is a LIST of child widgets (`Column.children`) —
    /// construction adds each element to the control. Single-child/scalar
    /// fields (`false`) are resolved per-value at construction runtime.
    pub children: bool,
}

impl FlutterField {
    pub(crate) const fn named(name: &'static str) -> Self {
        FlutterField { name, positional: None, default: None, children: false }
    }
    pub(crate) const fn named_default(name: &'static str, default: &'static str) -> Self {
        FlutterField { name, positional: None, default: Some(default), children: false }
    }
    pub(crate) const fn positional(name: &'static str, slot: u8) -> Self {
        FlutterField { name, positional: Some(slot), default: None, children: false }
    }
    /// A list-of-children field (`Column(children: [...])`).
    pub(crate) const fn children_list(name: &'static str) -> Self {
        FlutterField { name, positional: None, default: Some("const []"), children: true }
    }
}

pub(crate) const NO_FIELDS: &[FlutterField] = &[];
pub(crate) const NO_INTERFACES: &[&str] = &[];

/// Shared single-`child` field spec — used by several categories
/// (`IntrinsicHeight`, `RepaintBoundary`, `SliverToBoxAdapter`, …), so it
/// lives here rather than being duplicated per module.
pub(crate) const F_CHILD_ONLY: &[FlutterField] = &[FlutterField::named("child")];

use crate::emitter::widgets;

/// Every category's entries, concatenated once into a single lookup slice.
fn all_classes() -> &'static [FlutterClass] {
    static ALL: OnceLock<Vec<FlutterClass>> = OnceLock::new();
    ALL.get_or_init(|| {
        let mut v: Vec<FlutterClass> = Vec::new();
        for slice in widgets::ALL_CATEGORIES {
            v.extend_from_slice(slice);
        }
        v
    })
}

/// The complete Flutter adapter catalog.
pub fn flutter_classes() -> &'static [FlutterClass] {
    all_classes()
}

/// Field names that are `double` in Flutter wherever they appear on a widget.
/// Unambiguous only — deliberately EXCLUDES polymorphic names like `value`
/// (double on Slider, bool on Checkbox, generic on DropdownButton).
const DOUBLE_FIELD_NAMES: &[&str] = &[
    "width", "height", "left", "top", "right", "bottom", "elevation", "opacity",
    "fontSize", "letterSpacing", "wordSpacing", "dx", "dy", "widthFactor",
    "heightFactor", "aspectRatio", "thickness", "indent", "endIndent", "angle",
    "radius", "spacing", "runSpacing", "strokeWidth", "blurRadius", "spreadRadius",
    "scale", "textScaleFactor", "cacheExtent", "itemExtent", "minWidth", "maxWidth",
    "minHeight", "maxHeight", "toolbarHeight", "leadingWidth", "titleSpacing",
    "borderRadius", "minValue", "maxValue", "progress", "spaceRadius",
];

/// `(OwnerType, field)` pairs whose name is in `DOUBLE_FIELD_NAMES` but which
/// are NOT `double` on this particular widget — the name is polymorphic across
/// the catalog. `SafeArea.left`/`top`/`right`/`bottom` are `bool` edge-enable
/// flags, not offsets; leaving them seeded as `double` would render the bool
/// as a number (`false` → `0`).
const NOT_DOUBLE: &[(&str, &str)] = &[
    ("SafeArea", "left"),
    ("SafeArea", "top"),
    ("SafeArea", "right"),
    ("SafeArea", "bottom"),
];

/// Value-typed fields whose type is another value type — needed so a CHAINED
/// double read resolves (`Padding.padding` → `EdgeInsets`, then
/// `EdgeInsets.left` → `double`). `(OwnerType, field, fieldType)`.
const VALUE_FIELD_TYPES: &[(&str, &str, &str)] = &[
    ("Padding", "padding", "EdgeInsets"),
    ("Container", "padding", "EdgeInsets"),
    ("Container", "margin", "EdgeInsets"),
    ("AnimatedPadding", "padding", "EdgeInsets"),
    ("SliverPadding", "padding", "EdgeInsets"),
    ("ListView", "padding", "EdgeInsets"),
    ("Card", "margin", "EdgeInsets"),
    ("Offset", "dx", "double"),
    ("Offset", "dy", "double"),
    ("Size", "width", "double"),
    ("Size", "height", "double"),
    ("Radius", "x", "double"),
    ("Radius", "y", "double"),
    ("Slider", "value", "double"),
    ("Slider", "min", "double"),
    ("Slider", "max", "double"),
];

/// Static `(OwnerType, field) → fieldType` seed for the Dart frontend's type
/// tracker, so `double` fields render Dart-style (`10.0`) and chained value
/// reads resolve. Derived from the catalog (every widget's double-named
/// fields) plus the explicit value-type chains. Data lives HERE with the
/// catalog; the walker only consumes it.
pub fn field_type_seed() -> Vec<(&'static str, &'static str, &'static str)> {
    let mut out: Vec<(&'static str, &'static str, &'static str)> = Vec::new();
    for class in flutter_classes() {
        for field in class.fields {
            if DOUBLE_FIELD_NAMES.contains(&field.name)
                && !NOT_DOUBLE.contains(&(class.name, field.name))
            {
                out.push((class.name, field.name, "double"));
            }
        }
    }
    out.extend_from_slice(VALUE_FIELD_TYPES);
    out
}

/// Widget type names that realize as TRANSPARENT wrappers — the realizer
/// renders their child in their place. Derived from the catalog so the Dart
/// runtime never duplicates the list.
pub fn transparent_types() -> Vec<&'static str> {
    flutter_classes()
        .iter()
        .filter(|c| c.transparent)
        .map(|c| c.name)
        .collect()
}

/// Property keys the backing `vybe:gui` controls actually act on (the
/// `set_property` vocabulary: typed commands plus item population). Anything
/// else a Flutter field would emit — `opacity`, `elevation`, `padding`,
/// `message`, `clipBehavior` — has no control command behind it, so forwarding
/// it does nothing except stamp noise (an `EdgeInsets` even stringifies to
/// `[object]`). The realizer forwards only these.
pub const LIVE_PROPERTIES: &[&str] = &[
    "text",
    "flex",
    "enabled",
    "visible",
    "readonly",
    "value",
    "checked",
    "ischecked",
    "selected",
    "selectedindex",
    "clearitems",
    "additem",
];

/// The `is`/`instanceof` ancestry for `class`, self first: e.g. `Scaffold` →
/// `["Scaffold", "StatefulWidget", "Widget"]`. Stamped as the object's
/// `__types` array so `x is StatefulWidget` matches by membership.
pub fn ancestry(class: &FlutterClass) -> Vec<&'static str> {
    let all = flutter_classes();
    let mut chain = vec![class.name];
    let mut parent = class.parent;
    while let Some(p) = parent {
        chain.push(p);
        parent = all.iter().find(|c| c.name == p).and_then(|c| c.parent);
    }
    chain
}
