//! Flutter widget catalog.
//!
//! Flutter-shaped adapter surface. Each entry lowers a Flutter widget onto
//! an existing `vybe_widgets`/`vybe:gui` control (the same host constructors
//! the dotnet WinForms and plib VCL adapters use) — no Flutter-specific host
//! functions. The catalog is pure DATA: resolution logic lives in the shared
//! namespace resolver; construction/field-capture/`is`-checks ride the shared
//! class machinery (`vybe_emitter::classes`).
//!
//! Categories:
//! - `Widget`         — abstract base chain (no backing control), identity only.
//! - `widget_class!`  — concrete widget backed by a `vybe:gui` control ctor.
//! - value types      — `Color`, `EdgeInsets`, … : plain data classes, no control.
//! - enums            — `Axis`, `MainAxisAlignment`, … : registered via the
//!                      shared enum foundation, identity-compared with `==`.

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
    const fn named(name: &'static str) -> Self {
        FlutterField { name, positional: None, default: None, children: false }
    }
    const fn named_default(name: &'static str, default: &'static str) -> Self {
        FlutterField { name, positional: None, default: Some(default), children: false }
    }
    const fn positional(name: &'static str, slot: u8) -> Self {
        FlutterField { name, positional: Some(slot), default: None, children: false }
    }
    /// A list-of-children field (`Column(children: [...])`).
    const fn children_list(name: &'static str) -> Self {
        FlutterField { name, positional: None, default: Some("const []"), children: true }
    }
}

const NO_FIELDS: &[FlutterField] = &[];
const NO_INTERFACES: &[&str] = &[];

/// The complete Flutter adapter catalog.
pub fn flutter_classes() -> &'static [FlutterClass] {
    CLASSES
}

/// The `is`/`instanceof` ancestry for `class`, self first: e.g. `Scaffold` →
/// `["Scaffold", "StatefulWidget", "Widget"]`. Stamped as the object's
/// `__types` array so `x is StatefulWidget` matches by membership.
pub fn ancestry(class: &FlutterClass) -> Vec<&'static str> {
    let mut chain = vec![class.name];
    let mut parent = class.parent;
    while let Some(p) = parent {
        chain.push(p);
        parent = CLASSES
            .iter()
            .find(|c| c.name == p)
            .and_then(|c| c.parent);
    }
    chain
}

macro_rules! abstract_class {
    ($name:literal, $parent:expr) => {
        FlutterClass {
            name: $name,
            parent: $parent,
            interfaces: NO_INTERFACES,
            fields: NO_FIELDS,
            widget_host_fn: None,
        }
    };
}

macro_rules! widget_class {
    ($name:literal, $parent:literal, $host:literal, $fields:expr) => {
        FlutterClass {
            name: $name,
            parent: Some($parent),
            interfaces: NO_INTERFACES,
            fields: $fields,
            widget_host_fn: Some($host),
        }
    };
}

/// A widget backed by a `vybe:gui` control BUT whose construction/field-capture
/// is all that most tests exercise — same as `widget_class!` but spelled out
/// where the parent is itself a concrete widget.
macro_rules! data_class {
    // Pure data / value type (no backing control): Color, Offset, FocusNode…
    ($name:literal, $parent:expr, $fields:expr) => {
        FlutterClass {
            name: $name,
            parent: $parent,
            interfaces: NO_INTERFACES,
            fields: $fields,
            widget_host_fn: None,
        }
    };
}

// Flex-family fields (Column/Row/Flex share these). `children` defaults to an
// empty list; enum-valued fields (direction/*Alignment/*Size) default to null
// until the enum surface lands — a null default is inert, whereas an
// `Axis.vertical` default would evaluate an undefined enum on omission.
const FLEX_FIELDS: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("mainAxisAlignment"),
    FlutterField::named("mainAxisSize"),
    FlutterField::named("crossAxisAlignment"),
    FlutterField::named("textDirection"),
    FlutterField::named("verticalDirection"),
    FlutterField::named("textBaseline"),
    FlutterField::named("clipBehavior"),
    FlutterField::named("direction"),
];

const SCAFFOLD_FIELDS: &[FlutterField] = &[
    FlutterField::named("appBar"),
    FlutterField::named("body"),
    FlutterField::named("floatingActionButton"),
    FlutterField::named("drawer"),
    FlutterField::named("bottomNavigationBar"),
    FlutterField::named("backgroundColor"),
];

const APPBAR_FIELDS: &[FlutterField] = &[
    FlutterField::named("title"),
    FlutterField::children_list("actions"),
    FlutterField::named("leading"),
    FlutterField::named("bottom"),
    FlutterField::named("elevation"),
    FlutterField::named("backgroundColor"),
    FlutterField::named("centerTitle"),
];

const TEXT_FIELDS: &[FlutterField] = &[
    FlutterField::positional("data", 0),
    FlutterField::named("style"),
    FlutterField::named("textAlign"),
    FlutterField::named("textDirection"),
    FlutterField::named_default("softWrap", "true"),
    FlutterField::named("overflow"),
    FlutterField::named("maxLines"),
    FlutterField::named("textSpan"),
];

const CONTAINER_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("alignment"),
    FlutterField::named("color"),
    FlutterField::named("constraints"),
    FlutterField::named("decoration"),
    FlutterField::named("foregroundDecoration"),
    FlutterField::named("height"),
    FlutterField::named("width"),
    FlutterField::named("margin"),
    FlutterField::named("padding"),
    FlutterField::named("transform"),
    FlutterField::named("clipBehavior"),
];

const STACK_FIELDS: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("alignment"),
    FlutterField::named("fit"),
    FlutterField::named("clipBehavior"),
    FlutterField::named("textDirection"),
];

const ALIGN_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("alignment"),
    FlutterField::named("heightFactor"),
    FlutterField::named("widthFactor"),
];

const CENTER_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("heightFactor"),
    FlutterField::named("widthFactor"),
];

const PADDING_FIELDS: &[FlutterField] = &[
    FlutterField::named("padding"),
    FlutterField::named("child"),
];

const SIZEDBOX_FIELDS: &[FlutterField] = &[
    FlutterField::named("width"),
    FlutterField::named("height"),
    FlutterField::named("child"),
];

const ICON_FIELDS: &[FlutterField] = &[
    FlutterField::positional("icon", 0),
    FlutterField::named("color"),
    FlutterField::named("size"),
    FlutterField::named("semanticLabel"),
    FlutterField::named("textDirection"),
];

const EXPANDED_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named_default("flex", "1"),
];

const FLEXIBLE_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named_default("flex", "1"),
    FlutterField::named("fit"),
];

const MATERIALAPP_FIELDS: &[FlutterField] = &[
    FlutterField::named("home"),
    FlutterField::named("title"),
    FlutterField::named("theme"),
    FlutterField::named("initialRoute"),
    FlutterField::named("routes"),
    FlutterField::named("color"),
    FlutterField::named_default("debugShowCheckedModeBanner", "true"),
];

const ELEVATEDBUTTON_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("onPressed"),
    FlutterField::named("onLongPress"),
    FlutterField::named("style"),
    FlutterField::named("focusNode"),
    FlutterField::named("icon"),
    FlutterField::named("label"),
    FlutterField::named("autofocus"),
];

const POSITIONED_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("left"),
    FlutterField::named("top"),
    FlutterField::named("right"),
    FlutterField::named("bottom"),
    FlutterField::named("width"),
    FlutterField::named("height"),
];

const RADIO_FIELDS: &[FlutterField] = &[
    FlutterField::named("value"),
    FlutterField::named("groupValue"),
    FlutterField::named("onChanged"),
    FlutterField::named("activeColor"),
    FlutterField::named("toggleable"),
    FlutterField::named("focusNode"),
];

// `Color(0xFF00FF00).value` — the ARGB int is the sole positional field.
const COLOR_FIELDS: &[FlutterField] = &[FlutterField::positional("value", 0)];

// ── Field specs for the mapping-bucket widget/value catalog ──────────────
const F_FITTEDBOX: &[FlutterField] = &[FlutterField::named("fit"), FlutterField::named("alignment"), FlutterField::named("clipBehavior"), FlutterField::named("child")];
const F_FRACTRANS: &[FlutterField] = &[FlutterField::named("translation"), FlutterField::named("transformHitTests"), FlutterField::named("child")];
const F_FRACSIZED: &[FlutterField] = &[FlutterField::named("widthFactor"), FlutterField::named("heightFactor"), FlutterField::named("alignment"), FlutterField::named("child")];
const F_FUTUREB: &[FlutterField] = &[FlutterField::named("future"), FlutterField::named("initialData"), FlutterField::named("builder")];
const F_GESTURE: &[FlutterField] = &[FlutterField::named("onTap"), FlutterField::named("onDoubleTap"), FlutterField::named("onLongPress"), FlutterField::named("behavior"), FlutterField::named("child")];
const F_GRIDVIEW: &[FlutterField] = &[FlutterField::children_list("children"), FlutterField::named("gridDelegate"), FlutterField::named("crossAxisCount"), FlutterField::named("maxCrossAxisExtent"), FlutterField::named("itemBuilder"), FlutterField::named("childrenDelegate"), FlutterField::named("scrollDirection")];
const F_HERO: &[FlutterField] = &[FlutterField::named("tag"), FlutterField::named("child"), FlutterField::named("transitionOnUserGestures"), FlutterField::named("flightShuttleBuilder"), FlutterField::named("placeholderBuilder")];
const F_HEROMODE: &[FlutterField] = &[FlutterField::named("enabled"), FlutterField::named("child")];
const F_ICONTHEME: &[FlutterField] = &[FlutterField::named("data"), FlutterField::named("child")];
const F_IMAGE: &[FlutterField] = &[FlutterField::named("image"), FlutterField::named("width"), FlutterField::named("height"), FlutterField::named("fit"), FlutterField::named("alignment"), FlutterField::named("repeat"), FlutterField::named("color"), FlutterField::named("colorBlendMode"), FlutterField::named("filterQuality")];
const F_CHILD_ONLY: &[FlutterField] = &[FlutterField::named("child")];
const F_INHNOTIFIER: &[FlutterField] = &[FlutterField::named("notifier"), FlutterField::named("child")];
const F_INTERACTIVE: &[FlutterField] = &[FlutterField::named("clipBehavior"), FlutterField::named("panEnabled"), FlutterField::named("scaleEnabled"), FlutterField::named("minScale"), FlutterField::named("maxScale"), FlutterField::named("constrained"), FlutterField::named("boundaryMargin"), FlutterField::named("child")];
const F_INTRINSICW: &[FlutterField] = &[FlutterField::named("stepWidth"), FlutterField::named("stepHeight"), FlutterField::named("child")];
const F_LISTVIEW: &[FlutterField] = &[FlutterField::children_list("children"), FlutterField::named("scrollDirection"), FlutterField::named("reverse"), FlutterField::named("itemCount"), FlutterField::named("itemBuilder"), FlutterField::named("separatorBuilder"), FlutterField::named("childrenDelegate")];
const F_WRAP: &[FlutterField] = &[FlutterField::children_list("children"), FlutterField::named("direction"), FlutterField::named("alignment"), FlutterField::named("spacing"), FlutterField::named("runAlignment"), FlutterField::named("runSpacing"), FlutterField::named("crossAxisAlignment"), FlutterField::named("textDirection"), FlutterField::named("verticalDirection"), FlutterField::named("clipBehavior")];
const F_TRANSFORM: &[FlutterField] = &[FlutterField::named("transform"), FlutterField::named("origin"), FlutterField::named("alignment"), FlutterField::named("transformHitTests"), FlutterField::named("angle"), FlutterField::named("offset"), FlutterField::named("scale"), FlutterField::named("scaleX"), FlutterField::named("scaleY"), FlutterField::named("child")];
const F_INDEXEDSTACK: &[FlutterField] = &[FlutterField::children_list("children"), FlutterField::named("index"), FlutterField::named("alignment"), FlutterField::named("sizing"), FlutterField::named("textDirection")];
const F_STATEFULB: &[FlutterField] = &[FlutterField::named("builder")];
const F_STREAMB: &[FlutterField] = &[FlutterField::named("stream"), FlutterField::named("builder"), FlutterField::named("initialData")];
const F_VLBUILDER: &[FlutterField] = &[FlutterField::named("valueListenable"), FlutterField::named("builder"), FlutterField::named("child")];
const F_SLIVERGRID: &[FlutterField] = &[FlutterField::children_list("children"), FlutterField::named("delegate"), FlutterField::named("gridDelegate"), FlutterField::named("crossAxisCount"), FlutterField::named("maxCrossAxisExtent"), FlutterField::named("itemCount"), FlutterField::named("itemBuilder")];
const F_SLIVERLIST: &[FlutterField] = &[FlutterField::named("delegate"), FlutterField::named("itemCount"), FlutterField::named("itemBuilder"), FlutterField::named("separatorBuilder")];
const F_SLIVERPAD: &[FlutterField] = &[FlutterField::named("padding"), FlutterField::named("sliver")];
const F_SPACER: &[FlutterField] = &[FlutterField::named_default("flex", "1")];
const F_FOCUSNODE: &[FlutterField] = &[
    FlutterField::named("debugLabel"),
    FlutterField::named_default("hasFocus", "false"),
    FlutterField::named_default("hasPrimaryFocus", "false"),
    FlutterField::named_default("canRequestFocus", "true"),
    FlutterField::named_default("skipTraversal", "false"),
    FlutterField::named_default("descendantsAreFocusable", "true"),
];
const F_RECOGNIZER: &[FlutterField] =
    &[FlutterField::named("debugOwner"), FlutterField::named("onTap")];
const F_DEFTEXTSTYLE: &[FlutterField] = &[FlutterField::named("style"), FlutterField::named("child")];

// Simple positional/named value types (no computed getters, no control).
const F_OFFSET: &[FlutterField] = &[FlutterField::positional("dx", 0), FlutterField::positional("dy", 1)];
const F_SIZE: &[FlutterField] = &[FlutterField::positional("width", 0), FlutterField::positional("height", 1)];
const F_ICONDATA: &[FlutterField] = &[FlutterField::positional("codePoint", 0), FlutterField::named("fontFamily")];
const F_ICONTHEMEDATA: &[FlutterField] = &[FlutterField::named("color"), FlutterField::named("size")];
const F_VALKEY: &[FlutterField] = &[FlutterField::positional("value", 0)];
const F_VALNOTIFIER: &[FlutterField] = &[FlutterField::positional("value", 0)];
const F_NETIMAGE: &[FlutterField] = &[FlutterField::positional("url", 0), FlutterField::named("scale")];
const F_ASSETIMAGE: &[FlutterField] = &[FlutterField::positional("assetName", 0), FlutterField::named("scale")];
const F_TEXTSTYLE: &[FlutterField] = &[FlutterField::named("fontSize"), FlutterField::named("color"), FlutterField::named("fontWeight"), FlutterField::named("fontFamily")];

// grp_ab material widgets
const F_FORM: &[FlutterField] = &[FlutterField::named("child"), FlutterField::named("autovalidateMode"), FlutterField::named("onChanged"), FlutterField::named("canPop"), FlutterField::named("onPopInvokedWithResult")];
const F_GRIDTILE: &[FlutterField] = &[FlutterField::named("child"), FlutterField::named("header"), FlutterField::named("footer")];
const F_GRIDTILEBAR: &[FlutterField] = &[FlutterField::named("backgroundColor"), FlutterField::named("leading"), FlutterField::named("title"), FlutterField::named("subtitle"), FlutterField::named("trailing")];
const F_ICONBUTTON: &[FlutterField] = &[FlutterField::named("icon"), FlutterField::named("onPressed"), FlutterField::named("iconSize"), FlutterField::named("color"), FlutterField::named("tooltip")];
const F_INKWELL: &[FlutterField] = &[FlutterField::named("child"), FlutterField::named("onTap"), FlutterField::named("onDoubleTap"), FlutterField::named("onLongPress"), FlutterField::named("splashColor"), FlutterField::named("highlightColor"), FlutterField::named("borderRadius")];
const F_INPUTDECORATOR: &[FlutterField] = &[FlutterField::named("decoration"), FlutterField::named("child"), FlutterField::named("baseStyle"), FlutterField::named("isFocused"), FlutterField::named("isHovering"), FlutterField::named("expands")];
const F_LINEARPROG: &[FlutterField] = &[FlutterField::named("value"), FlutterField::named("backgroundColor"), FlutterField::named("color"), FlutterField::named("minHeight"), FlutterField::named("valueColor")];
const F_LISTTILE: &[FlutterField] = &[FlutterField::named("title"), FlutterField::named("subtitle"), FlutterField::named("leading"), FlutterField::named("trailing"), FlutterField::named("isThreeLine"), FlutterField::named("dense"), FlutterField::named("onTap")];
const F_OUTLINEDBTN: &[FlutterField] = &[FlutterField::named("onPressed"), FlutterField::named("child"), FlutterField::named("enabled"), FlutterField::named("style"), FlutterField::named("focusNode"), FlutterField::named("icon"), FlutterField::named("label")];
const F_POPUPMENUBTN: &[FlutterField] = &[FlutterField::named("itemBuilder"), FlutterField::named("initialValue"), FlutterField::named("onSelected"), FlutterField::named("icon")];
const F_POPUPMENUITEM: &[FlutterField] = &[FlutterField::named("value"), FlutterField::named("child")];
const F_SLIDER: &[FlutterField] = &[FlutterField::named("value"), FlutterField::named("onChanged"), FlutterField::named("min"), FlutterField::named("max"), FlutterField::named("divisions"), FlutterField::named("activeColor"), FlutterField::named("inactiveColor")];
const F_SLIVERAPPBAR: &[FlutterField] = &[FlutterField::named("title"), FlutterField::named("floating"), FlutterField::named("pinned"), FlutterField::named("snap"), FlutterField::named("expandedHeight"), FlutterField::named("flexibleSpace")];
const F_STEPPER: &[FlutterField] = &[FlutterField::children_list("steps"), FlutterField::named("currentStep"), FlutterField::named("type")];
const F_SWITCH: &[FlutterField] = &[FlutterField::named("value"), FlutterField::named("onChanged"), FlutterField::named("activeColor"), FlutterField::named("activeTrackColor"), FlutterField::named("inactiveThumbColor"), FlutterField::named("inactiveTrackColor")];
const F_TABBARVIEW: &[FlutterField] = &[FlutterField::children_list("children"), FlutterField::named("physics"), FlutterField::named("dragStartBehavior"), FlutterField::named("viewportFraction")];
const F_TABBAR: &[FlutterField] = &[FlutterField::children_list("tabs"), FlutterField::named("isScrollable"), FlutterField::named("indicatorColor"), FlutterField::named("labelColor"), FlutterField::named("unselectedLabelColor"), FlutterField::named("indicatorWeight")];
const F_TEXTBUTTON: &[FlutterField] = &[FlutterField::named("onPressed"), FlutterField::named("child"), FlutterField::named("enabled"), FlutterField::named("style"), FlutterField::named("autofocus"), FlutterField::named("icon"), FlutterField::named("label")];
const F_TEXTFIELD: &[FlutterField] = &[FlutterField::named("controller"), FlutterField::named("focusNode"), FlutterField::named("decoration"), FlutterField::named("keyboardType"), FlutterField::named("obscureText"), FlutterField::named("maxLines"), FlutterField::named("onChanged")];
const F_TEXTFORMFIELD: &[FlutterField] = &[FlutterField::named("controller"), FlutterField::named("initialValue"), FlutterField::named("validator"), FlutterField::named("onSaved"), FlutterField::named("decoration"), FlutterField::named("obscureText")];
const F_FAB: &[FlutterField] = &[FlutterField::named("onPressed"), FlutterField::named("child"), FlutterField::named("tooltip"), FlutterField::named("backgroundColor")];
const F_BOTTOMNAV: &[FlutterField] = &[FlutterField::children_list("items"), FlutterField::named("currentIndex"), FlutterField::named("onTap"), FlutterField::named("type")];
const F_TAB: &[FlutterField] = &[FlutterField::named("text"), FlutterField::named("icon"), FlutterField::named("child")];
const F_FLEXSPACEBAR: &[FlutterField] = &[FlutterField::named("title"), FlutterField::named("background"), FlutterField::named("centerTitle")];

// grp_ab value types
const F_BORDERSIDE: &[FlutterField] = &[FlutterField::named("color"), FlutterField::named("width")];
const F_INPUTDECORATION: &[FlutterField] = &[FlutterField::named("labelText"), FlutterField::named("hintText"), FlutterField::named("border"), FlutterField::named("icon")];
const F_TEXTEDITCTRL: &[FlutterField] = &[FlutterField::named("text")];
const F_BOTTOMNAVITEM: &[FlutterField] = &[FlutterField::named("icon"), FlutterField::named("label")];
const F_STEP: &[FlutterField] = &[FlutterField::named("title"), FlutterField::named("subtitle"), FlutterField::named("content"), FlutterField::named("isActive"), FlutterField::named("state")];
const F_ALWAYSANIM: &[FlutterField] = &[FlutterField::positional("value", 0)];

// grp_aa material + cupertino widgets
const F_CUPBUTTON: &[FlutterField] = &[FlutterField::named("onPressed"), FlutterField::named("child"), FlutterField::named("color"), FlutterField::named("disabledColor")];
const F_CUPSLIDER: &[FlutterField] = &[FlutterField::named("value"), FlutterField::named("onChanged"), FlutterField::named("min"), FlutterField::named("max"), FlutterField::named("divisions"), FlutterField::named("activeColor"), FlutterField::named("thumbColor")];
const F_CUPSWITCH: &[FlutterField] = &[FlutterField::named("value"), FlutterField::named("onChanged"), FlutterField::named("activeColor"), FlutterField::named("trackColor"), FlutterField::named("thumbColor")];
const F_BOTTOMAPPBAR: &[FlutterField] = &[FlutterField::named("color"), FlutterField::named("elevation"), FlutterField::named("shape"), FlutterField::named("clipBehavior"), FlutterField::named("notchMargin"), FlutterField::named("child")];
const F_BOTTOMSHEET: &[FlutterField] = &[FlutterField::named("onClosing"), FlutterField::named("builder"), FlutterField::named("elevation"), FlutterField::named("enableDrag"), FlutterField::named("onDragStart"), FlutterField::named("animationController")];
const F_CARD: &[FlutterField] = &[FlutterField::named("child"), FlutterField::named("color"), FlutterField::named("elevation"), FlutterField::named("shape"), FlutterField::named("margin"), FlutterField::named("clipBehavior")];
const F_CHECKBOX: &[FlutterField] = &[FlutterField::named("value"), FlutterField::named("onChanged"), FlutterField::named("tristate"), FlutterField::named("activeColor"), FlutterField::named("checkColor"), FlutterField::named("isError"), FlutterField::named("focusNode")];
const F_CHIP: &[FlutterField] = &[FlutterField::named("label"), FlutterField::named("avatar"), FlutterField::named("deleteIcon"), FlutterField::named("onDeleted")];
const F_ACTIONCHIP: &[FlutterField] = &[FlutterField::named("label"), FlutterField::named("onPressed")];
const F_SELCHIP: &[FlutterField] = &[FlutterField::named("label"), FlutterField::named("selected"), FlutterField::named("onSelected")];
const F_CIRCPROG: &[FlutterField] = &[FlutterField::named("value"), FlutterField::named("backgroundColor"), FlutterField::named("color"), FlutterField::named("strokeWidth"), FlutterField::named("strokeAlign")];
const F_DATATABLE: &[FlutterField] = &[FlutterField::children_list("columns"), FlutterField::children_list("rows"), FlutterField::named("sortColumnIndex"), FlutterField::named("sortAscending")];
const F_DATEPICKERDIALOG: &[FlutterField] = &[FlutterField::named("initialDate"), FlutterField::named("firstDate"), FlutterField::named("lastDate"), FlutterField::named("helpText")];
const F_DIVIDER: &[FlutterField] = &[FlutterField::named("height"), FlutterField::named("thickness"), FlutterField::named("indent"), FlutterField::named("endIndent"), FlutterField::named("color")];
const F_VDIVIDER: &[FlutterField] = &[FlutterField::named("width"), FlutterField::named("thickness"), FlutterField::named("indent"), FlutterField::named("endIndent"), FlutterField::named("color")];
const F_DRAWER: &[FlutterField] = &[FlutterField::named("child"), FlutterField::named("elevation"), FlutterField::named("semanticLabel")];
const F_DRAWERHEADER: &[FlutterField] = &[FlutterField::named("child"), FlutterField::named("margin"), FlutterField::named("decoration")];
const F_DROPDOWNBTN: &[FlutterField] = &[FlutterField::children_list("items"), FlutterField::named("onChanged"), FlutterField::named("value"), FlutterField::named("icon"), FlutterField::named("isExpanded")];
const F_DROPDOWNITEM: &[FlutterField] = &[FlutterField::named("value"), FlutterField::named("child")];
const F_CIRCLEAVATAR: &[FlutterField] = &[FlutterField::named("child"), FlutterField::named("radius"), FlutterField::named("backgroundColor"), FlutterField::named("backgroundImage")];

// grp_ae widgets
const F_OPACITY: &[FlutterField] = &[FlutterField::named("opacity"), FlutterField::named("alwaysIncludeSemantics"), FlutterField::named("child")];
const F_ANIMOPACITY: &[FlutterField] = &[FlutterField::named("opacity"), FlutterField::named("duration"), FlutterField::named("curve"), FlutterField::named("child")];
const F_SLIVEROPACITY: &[FlutterField] = &[FlutterField::named("opacity"), FlutterField::named("alwaysIncludeSemantics"), FlutterField::named("sliver")];
const F_ANIMPADDING: &[FlutterField] = &[FlutterField::named("padding"), FlutterField::named("duration"), FlutterField::named("curve"), FlutterField::named("child")];
const F_ANIMSIZE: &[FlutterField] = &[FlutterField::named("duration"), FlutterField::named("curve"), FlutterField::named("child")];
const F_POSDIR: &[FlutterField] = &[FlutterField::named("start"), FlutterField::named("end"), FlutterField::named("top"), FlutterField::named("bottom"), FlutterField::named("child")];
const F_RICHTEXT: &[FlutterField] = &[FlutterField::named("text"), FlutterField::named("textAlign"), FlutterField::named("textDirection"), FlutterField::named("softWrap"), FlutterField::named("overflow"), FlutterField::named("maxLines")];
const F_ROTATEDBOX: &[FlutterField] = &[FlutterField::named("quarterTurns"), FlutterField::named("child")];
const F_SAFEAREA: &[FlutterField] = &[FlutterField::named("left"), FlutterField::named("top"), FlutterField::named("right"), FlutterField::named("bottom"), FlutterField::named("minimum"), FlutterField::named("maintainBottomViewPadding"), FlutterField::named("child")];
const F_SLIVERSAFEAREA: &[FlutterField] = &[FlutterField::named("sliver"), FlutterField::named("minimum")];
const F_SCROLLBAR: &[FlutterField] = &[FlutterField::named("child"), FlutterField::named("controller"), FlutterField::named("thumbVisibility"), FlutterField::named("trackVisibility"), FlutterField::named("thickness"), FlutterField::named("radius")];
const F_SHADERMASK: &[FlutterField] = &[FlutterField::named("shaderCallback"), FlutterField::named("blendMode"), FlutterField::named("child")];
const F_SINGLESCROLL: &[FlutterField] = &[FlutterField::named("child"), FlutterField::named("scrollDirection"), FlutterField::named("reverse"), FlutterField::named("padding"), FlutterField::named("primary"), FlutterField::named("physics")];
const F_PAGEVIEW: &[FlutterField] = &[FlutterField::children_list("children"), FlutterField::named("scrollDirection"), FlutterField::named("reverse"), FlutterField::named("controller"), FlutterField::named("physics"), FlutterField::named("pageSnapping"), FlutterField::named("onPageChanged"), FlutterField::named("itemCount"), FlutterField::named("itemBuilder"), FlutterField::named("childrenDelegate")];
const F_PHYSMODEL: &[FlutterField] = &[FlutterField::named("color"), FlutterField::named("child"), FlutterField::named("shape"), FlutterField::named("clipBehavior"), FlutterField::named("elevation"), FlutterField::named("shadowColor")];
const F_PHYSSHAPE: &[FlutterField] = &[FlutterField::named("clipper"), FlutterField::named("color"), FlutterField::named("child"), FlutterField::named("clipBehavior"), FlutterField::named("elevation"), FlutterField::named("shadowColor")];

// Value types (grp_aa/ae/af) — unnamed ctor with named/positional args, no
// computed getters. (EdgeInsets/Rect/Radius/BorderRadius/BoxConstraints/Matrix4
// use named constructors + computed getters — separate follow-up.)
const F_GLOBALKEY: &[FlutterField] = &[FlutterField::named("debugLabel")];
const F_DATACOLUMN: &[FlutterField] = &[FlutterField::named("label"), FlutterField::named("tooltip"), FlutterField::named("numeric")];
const F_DATAROW: &[FlutterField] = &[FlutterField::named("cells"), FlutterField::named("selected")];
const F_DATACELL: &[FlutterField] = &[FlutterField::positional("child", 0), FlutterField::named("showEditIcon")];
const F_ROUTESETTINGS: &[FlutterField] = &[FlutterField::named("name"), FlutterField::named("arguments")];
const F_PAGECONTROLLER: &[FlutterField] = &[FlutterField::named("initialPage"), FlutterField::named("keepPage"), FlutterField::named("viewportFraction")];
const F_SCROLLCONTROLLER: &[FlutterField] = &[FlutterField::named("initialScrollOffset"), FlutterField::named("keepScrollOffset"), FlutterField::named("debugLabel")];
const F_FIXEDEXTENT: &[FlutterField] = &[FlutterField::named("initialItem")];
const F_SEMNODE: &[FlutterField] = &[FlutterField::named("label"), FlutterField::named("value"), FlutterField::named("hint"), FlutterField::named("increasedValue"), FlutterField::named("decreasedValue"), FlutterField::named("hasCheckedState"), FlutterField::named("isChecked"), FlutterField::named("rect"), FlutterField::named("tags")];
const F_LABEL_ONLY: &[FlutterField] = &[FlutterField::named("label")];
const F_SEMTAG: &[FlutterField] = &[FlutterField::positional("name", 0)];
const F_SEMACTION: &[FlutterField] = &[FlutterField::named("hint"), FlutterField::named("action")];
const F_SLIVERBUILDERDELEGATE: &[FlutterField] = &[FlutterField::positional("builder", 0), FlutterField::named("childCount"), FlutterField::named("findChildIndexCallback")];
const F_SLIVERLISTDELEGATE: &[FlutterField] = &[FlutterField::positional("children", 0)];
const F_LINEARGRAD: &[FlutterField] = &[FlutterField::named("colors"), FlutterField::named("begin"), FlutterField::named("end"), FlutterField::named("stops"), FlutterField::named("tileMode")];
const F_RADIALGRAD: &[FlutterField] = &[FlutterField::named("center"), FlutterField::named("radius"), FlutterField::named("colors"), FlutterField::named("tileMode")];
const F_TEXTSPAN2: &[FlutterField] = &[FlutterField::named("text"), FlutterField::named("children"), FlutterField::named("style")];
const F_BUTTONSTYLE: &[FlutterField] = &[FlutterField::named("backgroundColor"), FlutterField::named("elevation"), FlutterField::named("foregroundColor")];
const F_ROUNDEDBORDER: &[FlutterField] = &[FlutterField::named("borderRadius"), FlutterField::named("side")];
const F_ANIMCONTROLLER: &[FlutterField] = &[FlutterField::named("vsync"), FlutterField::named("duration"), FlutterField::named("lowerBound"), FlutterField::named("upperBound")];
const F_TICKER: &[FlutterField] = &[FlutterField::positional("onTick", 0)];
const F_IMGCONFIG: &[FlutterField] = &[FlutterField::named("size"), FlutterField::named("devicePixelRatio")];
const F_RESIZEIMAGE: &[FlutterField] = &[FlutterField::positional("imageProvider", 0), FlutterField::named("width"), FlutterField::named("height")];
const F_ASYNCSNAPSHOT: &[FlutterField] = &[FlutterField::named("connectionState"), FlutterField::named("hasData"), FlutterField::named("data"), FlutterField::named("hasError")];
const F_SGRIDDELEGATE: &[FlutterField] = &[FlutterField::named("crossAxisCount"), FlutterField::named("mainAxisSpacing"), FlutterField::named("crossAxisSpacing"), FlutterField::named("childAspectRatio")];

// grp_ac widgets
const F_TIMEPICKERDIALOG: &[FlutterField] = &[FlutterField::named("initialTime"), FlutterField::named("helpText"), FlutterField::named("cancelText"), FlutterField::named("confirmText")];
const F_TOOLTIP: &[FlutterField] = &[FlutterField::named("message"), FlutterField::named("child"), FlutterField::named("richMessage"), FlutterField::named("height"), FlutterField::named("padding"), FlutterField::named("waitDuration"), FlutterField::named("showDuration")];
const F_ANIMALIGN: &[FlutterField] = &[FlutterField::named("alignment"), FlutterField::named("duration"), FlutterField::named("child"), FlutterField::named("curve")];
const F_ANIMBUILDER: &[FlutterField] = &[FlutterField::named("animation"), FlutterField::named("listenable"), FlutterField::named("child"), FlutterField::named("builder")];
const F_ASPECTRATIO: &[FlutterField] = &[FlutterField::named("aspectRatio"), FlutterField::named("child")];
const F_BACKDROP: &[FlutterField] = &[FlutterField::named("filter"), FlutterField::named("blendMode"), FlutterField::named("child")];
const F_BASELINE: &[FlutterField] = &[FlutterField::named("baseline"), FlutterField::named("baselineType"), FlutterField::named("child")];
const F_CLIPPER: &[FlutterField] = &[FlutterField::named("clipBehavior"), FlutterField::named("child"), FlutterField::named("clipper")];
const F_CLIPRRECT: &[FlutterField] = &[FlutterField::named("borderRadius"), FlutterField::named("child"), FlutterField::named("clipBehavior"), FlutterField::named("clipper")];
const F_COLORFILTERED: &[FlutterField] = &[FlutterField::named("colorFilter"), FlutterField::named("child")];
const F_CUSTOMPAINT: &[FlutterField] = &[FlutterField::named("painter"), FlutterField::named("foregroundPainter"), FlutterField::named("size"), FlutterField::named("isComplex"), FlutterField::named("willChange"), FlutterField::named("child")];
const F_CUSTOMSCROLL: &[FlutterField] = &[FlutterField::children_list("slivers"), FlutterField::named("scrollDirection"), FlutterField::named("reverse"), FlutterField::named("primary"), FlutterField::named("physics"), FlutterField::named("anchor"), FlutterField::named("center")];
const F_DECORATEDBOX: &[FlutterField] = &[FlutterField::named("decoration"), FlutterField::named("position"), FlutterField::named("child")];
const F_DISMISSIBLE: &[FlutterField] = &[FlutterField::named("key"), FlutterField::named("child"), FlutterField::named("background"), FlutterField::named("secondaryBackground"), FlutterField::named("direction"), FlutterField::named("onDismissed")];
const F_DRAGGABLE: &[FlutterField] = &[FlutterField::named("data"), FlutterField::named("feedback"), FlutterField::named("child"), FlutterField::named("childWhenDragging")];
const F_DRAGTARGET: &[FlutterField] = &[FlutterField::named("builder"), FlutterField::named("onWillAcceptWithDetails"), FlutterField::named("onAcceptWithDetails")];

// Remaining value-type specs
const F_TIMEOFDAY: &[FlutterField] = &[FlutterField::named("hour"), FlutterField::named("minute")];
const F_BOXDECORATION: &[FlutterField] = &[FlutterField::named("color"), FlutterField::named("shape"), FlutterField::named("borderRadius"), FlutterField::named("border"), FlutterField::named("gradient")];
const F_FILEIMAGE: &[FlutterField] = &[FlutterField::positional("file", 0)];
const F_MEMIMAGE: &[FlutterField] = &[FlutterField::positional("bytes", 0)];
const F_POINTER: &[FlutterField] = &[FlutterField::named("pointer")];
const F_FITTEDSIZES: &[FlutterField] = &[FlutterField::positional("source", 0), FlutterField::positional("destination", 1)];
const F_FOCUSMANAGER: &[FlutterField] = &[FlutterField::named("rootScope")];
const F_FOCUSSCOPENODE: &[FlutterField] = &[FlutterField::named("debugLabel")];
const F_DIAGBUILDER: &[FlutterField] = &[FlutterField::named("properties")];

static CLASSES: &[FlutterClass] = &[
    // ── Abstract base chain — identity only, no backing control ──────────
    abstract_class!("Widget", None),
    abstract_class!("StatelessWidget", Some("Widget")),
    abstract_class!("StatefulWidget", Some("Widget")),
    abstract_class!("RenderObjectWidget", Some("Widget")),
    abstract_class!("SingleChildRenderObjectWidget", Some("RenderObjectWidget")),
    abstract_class!("MultiChildRenderObjectWidget", Some("RenderObjectWidget")),
    abstract_class!("ProxyWidget", Some("Widget")),
    abstract_class!("ParentDataWidget", Some("ProxyWidget")),
    // ── Concrete widgets (config objects; vybe_widgets backing at runApp) ─
    widget_class!("Scaffold", "StatefulWidget", "FlowLayoutPanel", SCAFFOLD_FIELDS),
    widget_class!("AppBar", "StatefulWidget", "FlowLayoutPanel", APPBAR_FIELDS),
    widget_class!("Text", "StatelessWidget", "Label", TEXT_FIELDS),
    widget_class!("Placeholder", "StatelessWidget", "Panel", NO_FIELDS),
    widget_class!("Container", "StatelessWidget", "FlowLayoutPanel", CONTAINER_FIELDS),
    widget_class!("Flex", "MultiChildRenderObjectWidget", "FlowLayoutPanel", FLEX_FIELDS),
    widget_class!("Column", "Flex", "FlowLayoutPanel", FLEX_FIELDS),
    widget_class!("Row", "Flex", "HFlowLayoutPanel", FLEX_FIELDS),
    widget_class!("Stack", "MultiChildRenderObjectWidget", "FlowLayoutPanel", STACK_FIELDS),
    widget_class!("Align", "SingleChildRenderObjectWidget", "FlowLayoutPanel", ALIGN_FIELDS),
    widget_class!("Center", "Align", "FlowLayoutPanel", CENTER_FIELDS),
    widget_class!("Padding", "SingleChildRenderObjectWidget", "FlowLayoutPanel", PADDING_FIELDS),
    widget_class!("SizedBox", "SingleChildRenderObjectWidget", "FlowLayoutPanel", SIZEDBOX_FIELDS),
    widget_class!("Icon", "StatelessWidget", "Label", ICON_FIELDS),
    widget_class!("Flexible", "ParentDataWidget", "FlowLayoutPanel", FLEXIBLE_FIELDS),
    widget_class!("Expanded", "Flexible", "FlowLayoutPanel", EXPANDED_FIELDS),
    widget_class!("Positioned", "ParentDataWidget", "Panel", POSITIONED_FIELDS),
    // ── Material app + button surface (samples) ──────────────────────────
    widget_class!("MaterialApp", "StatefulWidget", "FlowLayoutPanel", MATERIALAPP_FIELDS),
    widget_class!("ElevatedButton", "StatefulWidget", "Button", ELEVATEDBUTTON_FIELDS),
    widget_class!("Radio", "StatefulWidget", "RadioButton", RADIO_FIELDS),
    // ── Abstract bases needed for `is` chains ────────────────────────────
    abstract_class!("InheritedWidget", Some("ProxyWidget")),
    abstract_class!("PreferredSizeWidget", Some("Widget")),
    abstract_class!("ScrollView", Some("StatelessWidget")),
    abstract_class!("BoxScrollView", Some("ScrollView")),
    abstract_class!("FormField", Some("StatefulWidget")),
    abstract_class!("ImplicitlyAnimatedWidget", Some("StatefulWidget")),
    abstract_class!("AnimatedWidget", Some("StatefulWidget")),
    // `Key('x')` is a factory for `ValueKey<String>('x')` — construct it as a
    // value key: capture `value`, and carry ValueKey/LocalKey identity so
    // `is ValueKey`/`is LocalKey` match. Value equality via the allow-list.
    FlutterClass {
        name: "Key",
        parent: None,
        interfaces: &["ValueKey", "LocalKey"],
        fields: F_VALKEY,
        widget_host_fn: None,
    },
    abstract_class!("LocalKey", Some("Key")),
    // ── grp_ad widgets ───────────────────────────────────────────────────
    widget_class!("FittedBox", "SingleChildRenderObjectWidget", "Panel", F_FITTEDBOX),
    widget_class!("FractionalTranslation", "SingleChildRenderObjectWidget", "Panel", F_FRACTRANS),
    widget_class!("FractionallySizedBox", "SingleChildRenderObjectWidget", "Panel", F_FRACSIZED),
    widget_class!("FutureBuilder", "StatefulWidget", "Panel", F_FUTUREB),
    widget_class!("GestureDetector", "StatelessWidget", "Panel", F_GESTURE),
    widget_class!("GridView", "BoxScrollView", "FlowLayoutPanel", F_GRIDVIEW),
    widget_class!("Hero", "StatefulWidget", "Panel", F_HERO),
    widget_class!("HeroMode", "StatelessWidget", "Panel", F_HEROMODE),
    widget_class!("IconTheme", "InheritedWidget", "Panel", F_ICONTHEME),
    widget_class!("Image", "StatefulWidget", "picturebox", F_IMAGE),
    widget_class!("InheritedNotifier", "InheritedWidget", "Panel", F_INHNOTIFIER),
    widget_class!("InheritedModel", "InheritedWidget", "Panel", F_CHILD_ONLY),
    widget_class!("InteractiveViewer", "StatefulWidget", "Panel", F_INTERACTIVE),
    widget_class!("IntrinsicHeight", "SingleChildRenderObjectWidget", "Panel", F_CHILD_ONLY),
    widget_class!("IntrinsicWidth", "SingleChildRenderObjectWidget", "Panel", F_INTRINSICW),
    widget_class!("ListView", "BoxScrollView", "listbox", F_LISTVIEW),
    // ── grp_af widgets ───────────────────────────────────────────────────
    widget_class!("SliverGrid", "StatelessWidget", "FlowLayoutPanel", F_SLIVERGRID),
    widget_class!("SliverList", "StatelessWidget", "FlowLayoutPanel", F_SLIVERLIST),
    widget_class!("SliverPadding", "SingleChildRenderObjectWidget", "Panel", F_SLIVERPAD),
    widget_class!("SliverToBoxAdapter", "SingleChildRenderObjectWidget", "Panel", F_CHILD_ONLY),
    widget_class!("DefaultTextStyle", "StatelessWidget", "Panel", F_DEFTEXTSTYLE),
    widget_class!("Spacer", "StatelessWidget", "Panel", F_SPACER),
    widget_class!("IndexedStack", "Stack", "FlowLayoutPanel", F_INDEXEDSTACK),
    widget_class!("StatefulBuilder", "StatefulWidget", "Panel", F_STATEFULB),
    widget_class!("StreamBuilder", "StatefulWidget", "Panel", F_STREAMB),
    widget_class!("Transform", "SingleChildRenderObjectWidget", "Panel", F_TRANSFORM),
    widget_class!("ValueListenableBuilder", "StatefulWidget", "Panel", F_VLBUILDER),
    widget_class!("Wrap", "MultiChildRenderObjectWidget", "FlowLayoutPanel", F_WRAP),
    // ── grp_ab material widgets ──────────────────────────────────────────
    widget_class!("Form", "StatefulWidget", "Panel", F_FORM),
    widget_class!("GridTile", "StatelessWidget", "Panel", F_GRIDTILE),
    widget_class!("GridTileBar", "StatelessWidget", "Panel", F_GRIDTILEBAR),
    widget_class!("IconButton", "StatelessWidget", "Button", F_ICONBUTTON),
    widget_class!("InkWell", "StatelessWidget", "Panel", F_INKWELL),
    widget_class!("InputDecorator", "StatefulWidget", "Panel", F_INPUTDECORATOR),
    widget_class!("LinearProgressIndicator", "StatefulWidget", "progressbar", F_LINEARPROG),
    widget_class!("ListTile", "StatelessWidget", "Panel", F_LISTTILE),
    widget_class!("OutlinedButton", "StatefulWidget", "Button", F_OUTLINEDBTN),
    widget_class!("PopupMenuButton", "StatefulWidget", "Button", F_POPUPMENUBTN),
    widget_class!("PopupMenuItem", "Widget", "Panel", F_POPUPMENUITEM),
    widget_class!("Slider", "StatefulWidget", "trackbar", F_SLIDER),
    widget_class!("SliverAppBar", "StatefulWidget", "Panel", F_SLIVERAPPBAR),
    widget_class!("Stepper", "StatefulWidget", "Panel", F_STEPPER),
    widget_class!("Switch", "StatefulWidget", "CheckBox", F_SWITCH),
    widget_class!("TabBarView", "StatefulWidget", "FlowLayoutPanel", F_TABBARVIEW),
    widget_class!("TabBar", "StatefulWidget", "tabcontrol", F_TABBAR),
    widget_class!("TextButton", "StatefulWidget", "Button", F_TEXTBUTTON),
    widget_class!("TextField", "StatefulWidget", "TextBox", F_TEXTFIELD),
    widget_class!("TextFormField", "FormField", "TextBox", F_TEXTFORMFIELD),
    widget_class!("FloatingActionButton", "StatelessWidget", "Button", F_FAB),
    widget_class!("Drawer", "StatelessWidget", "Panel", F_DRAWER),
    widget_class!("BottomNavigationBar", "StatefulWidget", "HFlowLayoutPanel", F_BOTTOMNAV),
    widget_class!("FlexibleSpaceBar", "StatefulWidget", "Panel", F_FLEXSPACEBAR),
    widget_class!("Tab", "StatelessWidget", "Label", F_TAB),
    // ── grp_aa material + cupertino widgets ──────────────────────────────
    widget_class!("CupertinoButton", "StatefulWidget", "Button", F_CUPBUTTON),
    widget_class!("CupertinoSlider", "StatefulWidget", "trackbar", F_CUPSLIDER),
    widget_class!("CupertinoSwitch", "StatefulWidget", "CheckBox", F_CUPSWITCH),
    widget_class!("BottomAppBar", "StatefulWidget", "Panel", F_BOTTOMAPPBAR),
    widget_class!("BottomSheet", "StatefulWidget", "Panel", F_BOTTOMSHEET),
    widget_class!("Card", "StatelessWidget", "groupbox", F_CARD),
    widget_class!("Checkbox", "StatefulWidget", "CheckBox", F_CHECKBOX),
    widget_class!("Chip", "StatelessWidget", "Panel", F_CHIP),
    widget_class!("ActionChip", "StatelessWidget", "Button", F_ACTIONCHIP),
    widget_class!("FilterChip", "StatelessWidget", "Panel", F_SELCHIP),
    widget_class!("ChoiceChip", "StatelessWidget", "Panel", F_SELCHIP),
    widget_class!("CircularProgressIndicator", "StatefulWidget", "progressbar", F_CIRCPROG),
    widget_class!("DataTable", "StatelessWidget", "datagrid", F_DATATABLE),
    widget_class!("DatePickerDialog", "StatefulWidget", "datetimepicker", F_DATEPICKERDIALOG),
    widget_class!("Divider", "StatelessWidget", "Panel", F_DIVIDER),
    widget_class!("VerticalDivider", "StatelessWidget", "Panel", F_VDIVIDER),
    widget_class!("DrawerHeader", "StatelessWidget", "Panel", F_DRAWERHEADER),
    widget_class!("DropdownButton", "StatefulWidget", "combobox", F_DROPDOWNBTN),
    widget_class!("DropdownMenuItem", "Widget", "Panel", F_DROPDOWNITEM),
    widget_class!("BackButton", "StatelessWidget", "Button", NO_FIELDS),
    widget_class!("CircleAvatar", "StatelessWidget", "picturebox", F_CIRCLEAVATAR),
    // ── grp_ac widgets ───────────────────────────────────────────────────
    widget_class!("TimePickerDialog", "StatefulWidget", "datetimepicker", F_TIMEPICKERDIALOG),
    widget_class!("Tooltip", "StatefulWidget", "Panel", F_TOOLTIP),
    widget_class!("AnimatedAlign", "ImplicitlyAnimatedWidget", "Panel", F_ANIMALIGN),
    widget_class!("AnimatedBuilder", "AnimatedWidget", "Panel", F_ANIMBUILDER),
    widget_class!("AspectRatio", "SingleChildRenderObjectWidget", "Panel", F_ASPECTRATIO),
    widget_class!("BackdropFilter", "SingleChildRenderObjectWidget", "Panel", F_BACKDROP),
    widget_class!("Baseline", "SingleChildRenderObjectWidget", "Panel", F_BASELINE),
    widget_class!("ClipOval", "SingleChildRenderObjectWidget", "Panel", F_CLIPPER),
    widget_class!("ClipPath", "SingleChildRenderObjectWidget", "Panel", F_CLIPPER),
    widget_class!("ClipRect", "SingleChildRenderObjectWidget", "Panel", F_CLIPPER),
    widget_class!("ClipRRect", "SingleChildRenderObjectWidget", "Panel", F_CLIPRRECT),
    widget_class!("ColorFiltered", "SingleChildRenderObjectWidget", "Panel", F_COLORFILTERED),
    widget_class!("CustomPaint", "SingleChildRenderObjectWidget", "Panel", F_CUSTOMPAINT),
    widget_class!("CustomScrollView", "ScrollView", "FlowLayoutPanel", F_CUSTOMSCROLL),
    widget_class!("DecoratedBox", "SingleChildRenderObjectWidget", "Panel", F_DECORATEDBOX),
    widget_class!("Dismissible", "StatefulWidget", "Panel", F_DISMISSIBLE),
    widget_class!("Draggable", "StatefulWidget", "Panel", F_DRAGGABLE),
    widget_class!("DragTarget", "StatefulWidget", "Panel", F_DRAGTARGET),
    widget_class!("Opacity", "SingleChildRenderObjectWidget", "Panel", F_OPACITY),
    widget_class!("AnimatedOpacity", "ImplicitlyAnimatedWidget", "Panel", F_ANIMOPACITY),
    widget_class!("SliverOpacity", "SingleChildRenderObjectWidget", "Panel", F_SLIVEROPACITY),
    widget_class!("AnimatedPadding", "ImplicitlyAnimatedWidget", "Panel", F_ANIMPADDING),
    widget_class!("AnimatedSize", "StatefulWidget", "Panel", F_ANIMSIZE),
    widget_class!("PositionedDirectional", "ParentDataWidget", "Panel", F_POSDIR),
    widget_class!("RepaintBoundary", "SingleChildRenderObjectWidget", "Panel", F_CHILD_ONLY),
    widget_class!("RichText", "MultiChildRenderObjectWidget", "Label", F_RICHTEXT),
    widget_class!("RotatedBox", "SingleChildRenderObjectWidget", "Panel", F_ROTATEDBOX),
    widget_class!("SafeArea", "StatelessWidget", "Panel", F_SAFEAREA),
    widget_class!("SliverSafeArea", "StatelessWidget", "Panel", F_SLIVERSAFEAREA),
    widget_class!("Scrollbar", "StatelessWidget", "vscrollbar", F_SCROLLBAR),
    widget_class!("ShaderMask", "SingleChildRenderObjectWidget", "Panel", F_SHADERMASK),
    widget_class!("SingleChildScrollView", "StatelessWidget", "FlowLayoutPanel", F_SINGLESCROLL),
    widget_class!("PageView", "StatefulWidget", "FlowLayoutPanel", F_PAGEVIEW),
    widget_class!("PhysicalModel", "SingleChildRenderObjectWidget", "Panel", F_PHYSMODEL),
    widget_class!("PhysicalShape", "SingleChildRenderObjectWidget", "Panel", F_PHYSSHAPE),
    // ── Value types (no backing control; construction + field read-back) ──
    data_class!("Color", None, COLOR_FIELDS),
    data_class!("FocusNode", None, F_FOCUSNODE),
    data_class!("Offset", None, F_OFFSET),
    data_class!("Size", None, F_SIZE),
    data_class!("FractionalOffset", None, F_OFFSET),
    data_class!("IconData", None, F_ICONDATA),
    data_class!("IconThemeData", None, F_ICONTHEMEDATA),
    data_class!("TextStyle", None, F_TEXTSTYLE),
    data_class!("TextSpan", None, F_TEXTSPAN2),
    data_class!("BorderSide", None, F_BORDERSIDE),
    data_class!("InputDecoration", None, F_INPUTDECORATION),
    data_class!("TextEditingController", None, F_TEXTEDITCTRL),
    data_class!("BottomNavigationBarItem", None, F_BOTTOMNAVITEM),
    data_class!("Step", None, F_STEP),
    data_class!("AlwaysStoppedAnimation", None, F_ALWAYSANIM),
    data_class!("DataColumn", None, F_DATACOLUMN),
    data_class!("DataRow", None, F_DATAROW),
    data_class!("DataCell", None, F_DATACELL),
    data_class!("RouteSettings", None, F_ROUTESETTINGS),
    data_class!("PageController", None, F_PAGECONTROLLER),
    data_class!("ScrollController", None, F_SCROLLCONTROLLER),
    data_class!("TrackingScrollController", Some("ScrollController"), F_SCROLLCONTROLLER),
    data_class!("FixedExtentScrollController", Some("ScrollController"), F_FIXEDEXTENT),
    data_class!("SemanticsNode", None, F_SEMNODE),
    data_class!("SemanticsConfiguration", None, F_LABEL_ONLY),
    data_class!("SemanticsProperties", None, F_LABEL_ONLY),
    data_class!("CustomSemanticsAction", None, F_SEMACTION),
    data_class!("SemanticsTag", None, F_SEMTAG),
    data_class!("SliverChildBuilderDelegate", None, F_SLIVERBUILDERDELEGATE),
    data_class!("SliverChildListDelegate", None, F_SLIVERLISTDELEGATE),
    data_class!("SliverGridDelegateWithFixedCrossAxisCount", None, F_SGRIDDELEGATE),
    data_class!("LinearGradient", None, F_LINEARGRAD),
    data_class!("RadialGradient", None, F_RADIALGRAD),
    data_class!("ButtonStyle", None, F_BUTTONSTYLE),
    data_class!("RoundedRectangleBorder", None, F_ROUNDEDBORDER),
    data_class!("AnimationController", None, F_ANIMCONTROLLER),
    data_class!("Ticker", None, F_TICKER),
    data_class!("ImageConfiguration", None, F_IMGCONFIG),
    data_class!("AsyncSnapshot", None, F_ASYNCSNAPSHOT),
    data_class!("ThemeData", None, NO_FIELDS),
    data_class!("BouncingScrollPhysics", None, NO_FIELDS),
    data_class!("CircularNotchedRectangle", None, NO_FIELDS),
    data_class!("HeroController", None, NO_FIELDS),
    data_class!("FocusManager", None, F_FOCUSMANAGER),
    data_class!("FocusScopeNode", None, F_FOCUSSCOPENODE),
    data_class!("FittedSizes", None, F_FITTEDSIZES),
    data_class!("TimeOfDay", None, F_TIMEOFDAY),
    data_class!("BoxDecoration", None, F_BOXDECORATION),
    data_class!("DiagnosticPropertiesBuilder", None, F_DIAGBUILDER),
    data_class!("ValueNotifier", None, F_VALNOTIFIER),
    data_class!("NetworkImage", None, F_NETIMAGE),
    data_class!("AssetImage", None, F_ASSETIMAGE),
    data_class!("ExactAssetImage", None, F_ASSETIMAGE),
    data_class!("FileImage", None, F_FILEIMAGE),
    data_class!("MemoryImage", None, F_MEMIMAGE),
    data_class!("ResizeImage", None, F_RESIZEIMAGE),
    data_class!("PointerDownEvent", None, F_POINTER),
    data_class!("TapGestureRecognizer", None, F_RECOGNIZER),
    data_class!("DoubleTapGestureRecognizer", None, F_RECOGNIZER),
    data_class!("LongPressGestureRecognizer", None, F_RECOGNIZER),
    data_class!("PanGestureRecognizer", None, F_RECOGNIZER),
    data_class!("ScaleGestureRecognizer", None, F_RECOGNIZER),
    // ── Key hierarchy ────────────────────────────────────────────────────
    data_class!("ValueKey", Some("LocalKey"), F_VALKEY),
    data_class!("ObjectKey", Some("LocalKey"), F_VALKEY),
    data_class!("UniqueKey", Some("LocalKey"), NO_FIELDS),
    data_class!("GlobalKey", Some("Key"), F_GLOBALKEY),
    data_class!("GlobalObjectKey", Some("GlobalKey"), F_VALKEY),
];
