//! Scrolling widgets & slivers — ListView/GridView, CustomScrollView, the
//! `Sliver*` family, SafeArea, plus the scroll controllers and sliver child
//! delegates they take as data arguments.

use crate::emitter::catalog::{F_CHILD_ONLY, FlutterClass, FlutterField};

const F_LISTVIEW: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("scrollDirection"),
    FlutterField::named("reverse"),
    FlutterField::named("itemCount"),
    FlutterField::named("itemBuilder"),
    FlutterField::named("separatorBuilder"),
    FlutterField::named("childrenDelegate"),
];

const F_GRIDVIEW: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("gridDelegate"),
    FlutterField::named("crossAxisCount"),
    FlutterField::named("maxCrossAxisExtent"),
    FlutterField::named("itemBuilder"),
    FlutterField::named("childrenDelegate"),
    FlutterField::named("scrollDirection"),
];

const F_CUSTOMSCROLL: &[FlutterField] = &[
    FlutterField::children_list("slivers"),
    FlutterField::named("scrollDirection"),
    FlutterField::named("reverse"),
    FlutterField::named("primary"),
    FlutterField::named("physics"),
    FlutterField::named("anchor"),
    FlutterField::named("center"),
];

const F_SINGLESCROLL: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("scrollDirection"),
    FlutterField::named("reverse"),
    FlutterField::named("padding"),
    FlutterField::named("primary"),
    FlutterField::named("physics"),
];

const F_PAGEVIEW: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("scrollDirection"),
    FlutterField::named("reverse"),
    FlutterField::named("controller"),
    FlutterField::named("physics"),
    FlutterField::named("pageSnapping"),
    FlutterField::named("onPageChanged"),
    FlutterField::named("itemCount"),
    FlutterField::named("itemBuilder"),
    FlutterField::named("childrenDelegate"),
];

const F_SCROLLBAR: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("controller"),
    FlutterField::named("thumbVisibility"),
    FlutterField::named("trackVisibility"),
    FlutterField::named("thickness"),
    FlutterField::named("radius"),
];

const F_SLIVERGRID: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("delegate"),
    FlutterField::named("gridDelegate"),
    FlutterField::named("crossAxisCount"),
    FlutterField::named("maxCrossAxisExtent"),
    FlutterField::named("itemCount"),
    FlutterField::named("itemBuilder"),
];

const F_SLIVERLIST: &[FlutterField] = &[
    FlutterField::named("delegate"),
    FlutterField::named("itemCount"),
    FlutterField::named("itemBuilder"),
    FlutterField::named("separatorBuilder"),
];

const F_SLIVERPAD: &[FlutterField] = &[
    FlutterField::named("padding"),
    FlutterField::named("sliver"),
];

const F_SLIVERAPPBAR: &[FlutterField] = &[
    FlutterField::named("title"),
    FlutterField::named("floating"),
    FlutterField::named("pinned"),
    FlutterField::named("snap"),
    FlutterField::named("expandedHeight"),
    FlutterField::named("flexibleSpace"),
];

const F_SAFEAREA: &[FlutterField] = &[
    FlutterField::named("left"),
    FlutterField::named("top"),
    FlutterField::named("right"),
    FlutterField::named("bottom"),
    FlutterField::named("minimum"),
    FlutterField::named("maintainBottomViewPadding"),
    FlutterField::named("child"),
];

const F_SLIVERSAFEAREA: &[FlutterField] = &[
    FlutterField::named("sliver"),
    FlutterField::named("minimum"),
];

const F_SCROLLCONTROLLER: &[FlutterField] = &[
    FlutterField::named("initialScrollOffset"),
    FlutterField::named("keepScrollOffset"),
    FlutterField::named("debugLabel"),
];

const F_FIXEDEXTENT: &[FlutterField] = &[FlutterField::named("initialItem")];

const F_PAGECONTROLLER: &[FlutterField] = &[
    FlutterField::named("initialPage"),
    FlutterField::named("keepPage"),
    FlutterField::named("viewportFraction"),
];

const F_SLIVERBUILDERDELEGATE: &[FlutterField] = &[
    FlutterField::positional("builder", 0),
    FlutterField::named("childCount"),
    FlutterField::named("findChildIndexCallback"),
];

const F_SLIVERLISTDELEGATE: &[FlutterField] = &[FlutterField::positional("children", 0)];

const F_SGRIDDELEGATE: &[FlutterField] = &[
    FlutterField::named("crossAxisCount"),
    FlutterField::named("mainAxisSpacing"),
    FlutterField::named("crossAxisSpacing"),
    FlutterField::named("childAspectRatio"),
];

pub(crate) const CLASSES: &[FlutterClass] = &[
    // A scroll view is a column that overflows — `overflow: auto` is the whole
    // difference, and it is CSS, not a control kind.
    FlutterClass::widget(
        "ListView",
        "BoxScrollView",
        "div;display:flex;flex-direction:column;overflow:auto",
        F_LISTVIEW,
    ),
    // ⚠ NOT `display: grid` — grid parses and cascades in `vybe_widgets` with
    // no grid layout behind it, so naming it would be a silent no-op. Wrapping
    // flex is the honest approximation, and it is stated rather than implied.
    FlutterClass::widget(
        "GridView",
        "BoxScrollView",
        "div;display:flex;flex-wrap:wrap;overflow:auto",
        F_GRIDVIEW,
    ),
    FlutterClass::widget(
        "CustomScrollView",
        "ScrollView",
        "div;display:flex;flex-direction:column;overflow:auto",
        F_CUSTOMSCROLL,
    ),
    FlutterClass::widget(
        "SingleChildScrollView",
        "StatelessWidget",
        "div;overflow:auto",
        F_SINGLESCROLL,
    ),
    FlutterClass::widget(
        "PageView",
        "StatefulWidget",
        "div;display:flex;flex-direction:row;overflow:auto",
        F_PAGEVIEW,
    ),
    // A scrollbar is a value in a range, which `<input type=range>` IS — the
    // same element dotnet's `HScrollBar`/`VScrollBar` resolve to, turned onto
    // its side in CSS rather than by a second tag.
    FlutterClass::widget(
        "Scrollbar",
        "StatelessWidget",
        "input:range;writing-mode:vertical-lr",
        F_SCROLLBAR,
    ),
    FlutterClass::widget(
        "SliverGrid",
        "StatelessWidget",
        "div;display:flex;flex-wrap:wrap",
        F_SLIVERGRID,
    ),
    FlutterClass::widget(
        "SliverList",
        "StatelessWidget",
        "div;display:flex;flex-direction:column",
        F_SLIVERLIST,
    ),
    // Inset/adapter slivers only pad or re-box their sliver — no visual of
    // their own on the backing controls, so they realize transparently.
    FlutterClass::wrapper(
        "SliverPadding",
        "SingleChildRenderObjectWidget",
        F_SLIVERPAD,
    ),
    FlutterClass::wrapper(
        "SliverToBoxAdapter",
        "SingleChildRenderObjectWidget",
        F_CHILD_ONLY,
    ),
    FlutterClass::wrapper("SafeArea", "StatelessWidget", F_SAFEAREA),
    FlutterClass::wrapper("SliverSafeArea", "StatelessWidget", F_SLIVERSAFEAREA),
    FlutterClass::widget(
        "SliverAppBar",
        "StatefulWidget",
        "header;display:flex;align-items:center",
        F_SLIVERAPPBAR,
    ),
    FlutterClass::data("ScrollController", None, F_SCROLLCONTROLLER),
    FlutterClass::data(
        "TrackingScrollController",
        Some("ScrollController"),
        F_SCROLLCONTROLLER,
    ),
    FlutterClass::data(
        "FixedExtentScrollController",
        Some("ScrollController"),
        F_FIXEDEXTENT,
    ),
    FlutterClass::data("PageController", None, F_PAGECONTROLLER),
    FlutterClass::data("SliverChildBuilderDelegate", None, F_SLIVERBUILDERDELEGATE),
    FlutterClass::data("SliverChildListDelegate", None, F_SLIVERLISTDELEGATE),
    FlutterClass::data(
        "SliverGridDelegateWithFixedCrossAxisCount",
        None,
        F_SGRIDDELEGATE,
    ),
    FlutterClass::data("BouncingScrollPhysics", None, &[]),
];
