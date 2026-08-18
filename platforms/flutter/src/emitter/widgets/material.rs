//! Material widgets — app shell (Scaffold/AppBar/MaterialApp), buttons, cards,
//! lists, chips, dialogs, tabs, progress indicators, and their data types.

use crate::emitter::catalog::{FlutterClass, FlutterField};

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

const MATERIALAPP_FIELDS: &[FlutterField] = &[
    FlutterField::named("home"),
    // `MaterialApp.title` is the string the OS task switcher shows for the
    // application — the WINDOW title. On the web that is `document.title`, the
    // `<title>` element in the head. Left as the field's own name it became a
    // `title=""` attribute on the root div, which HTML defines as the hover
    // TOOLTIP: the app's name silently turned into markup nobody displays,
    // while the window kept a default caption.
    FlutterField::named_role("title", "windowtitle"),
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

const F_CARD: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("color"),
    FlutterField::named("elevation"),
    FlutterField::named("shape"),
    FlutterField::named("margin"),
    FlutterField::named("clipBehavior"),
];

const F_DRAWER: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("elevation"),
    FlutterField::named("semanticLabel"),
];

const F_DRAWERHEADER: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("margin"),
    FlutterField::named("decoration"),
];

const F_DIVIDER: &[FlutterField] = &[
    FlutterField::named("height"),
    FlutterField::named("thickness"),
    FlutterField::named("indent"),
    FlutterField::named("endIndent"),
    FlutterField::named("color"),
];

const F_VDIVIDER: &[FlutterField] = &[
    FlutterField::named("width"),
    FlutterField::named("thickness"),
    FlutterField::named("indent"),
    FlutterField::named("endIndent"),
    FlutterField::named("color"),
];

const F_LISTTILE: &[FlutterField] = &[
    FlutterField::named("title"),
    FlutterField::named("subtitle"),
    FlutterField::named("leading"),
    FlutterField::named("trailing"),
    FlutterField::named("isThreeLine"),
    FlutterField::named("dense"),
    FlutterField::named("onTap"),
];

const F_CHIP: &[FlutterField] = &[
    FlutterField::named("label"),
    FlutterField::named("avatar"),
    FlutterField::named("deleteIcon"),
    FlutterField::named("onDeleted"),
];

const F_ACTIONCHIP: &[FlutterField] = &[
    FlutterField::named("label"),
    FlutterField::named("onPressed"),
];

const F_SELCHIP: &[FlutterField] = &[
    FlutterField::named("label"),
    FlutterField::named("selected"),
    FlutterField::named("onSelected"),
];

const F_CIRCLEAVATAR: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("radius"),
    FlutterField::named("backgroundColor"),
    FlutterField::named("backgroundImage"),
];

const F_GRIDTILE: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("header"),
    FlutterField::named("footer"),
];

const F_GRIDTILEBAR: &[FlutterField] = &[
    FlutterField::named("backgroundColor"),
    FlutterField::named("leading"),
    FlutterField::named("title"),
    FlutterField::named("subtitle"),
    FlutterField::named("trailing"),
];

const F_STEPPER: &[FlutterField] = &[
    FlutterField::children_list("steps"),
    FlutterField::named("currentStep"),
    FlutterField::named("type"),
];

const F_STEP: &[FlutterField] = &[
    FlutterField::named("title"),
    FlutterField::named("subtitle"),
    FlutterField::named("content"),
    FlutterField::named("isActive"),
    FlutterField::named("state"),
];

const F_DATATABLE: &[FlutterField] = &[
    FlutterField::children_list("columns"),
    FlutterField::children_list("rows"),
    FlutterField::named("sortColumnIndex"),
    FlutterField::named("sortAscending"),
];

const F_DATACOLUMN: &[FlutterField] = &[
    FlutterField::named("label"),
    FlutterField::named("tooltip"),
    FlutterField::named("numeric"),
];

const F_DATAROW: &[FlutterField] = &[
    FlutterField::named("cells"),
    FlutterField::named("selected"),
];

const F_DATACELL: &[FlutterField] = &[
    FlutterField::positional("child", 0),
    FlutterField::named("showEditIcon"),
];

const F_BOTTOMNAV: &[FlutterField] = &[
    FlutterField::children_list("items"),
    FlutterField::named("currentIndex"),
    FlutterField::named("onTap"),
    FlutterField::named("type"),
];

const F_BOTTOMNAVITEM: &[FlutterField] =
    &[FlutterField::named("icon"), FlutterField::named("label")];

const F_BOTTOMAPPBAR: &[FlutterField] = &[
    FlutterField::named("color"),
    FlutterField::named("elevation"),
    FlutterField::named("shape"),
    FlutterField::named("clipBehavior"),
    FlutterField::named("notchMargin"),
    FlutterField::named("child"),
];

const F_BOTTOMSHEET: &[FlutterField] = &[
    FlutterField::named("onClosing"),
    FlutterField::named("builder"),
    FlutterField::named("elevation"),
    FlutterField::named("enableDrag"),
    FlutterField::named("onDragStart"),
    FlutterField::named("animationController"),
];

const F_TABBAR: &[FlutterField] = &[
    FlutterField::children_list("tabs"),
    FlutterField::named("isScrollable"),
    FlutterField::named("indicatorColor"),
    FlutterField::named("labelColor"),
    FlutterField::named("unselectedLabelColor"),
    FlutterField::named("indicatorWeight"),
];

const F_TABBARVIEW: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("physics"),
    FlutterField::named("dragStartBehavior"),
    FlutterField::named("viewportFraction"),
];

const F_TAB: &[FlutterField] = &[
    FlutterField::named("text"),
    FlutterField::named("icon"),
    FlutterField::named("child"),
];

const F_FLEXSPACEBAR: &[FlutterField] = &[
    FlutterField::named("title"),
    FlutterField::named("background"),
    FlutterField::named("centerTitle"),
];

const F_LINEARPROG: &[FlutterField] = &[
    FlutterField::named("value"),
    FlutterField::named("backgroundColor"),
    FlutterField::named("color"),
    FlutterField::named("minHeight"),
    FlutterField::named("valueColor"),
];

const F_CIRCPROG: &[FlutterField] = &[
    FlutterField::named("value"),
    FlutterField::named("backgroundColor"),
    FlutterField::named("color"),
    FlutterField::named("strokeWidth"),
    FlutterField::named("strokeAlign"),
];

const F_DATEPICKERDIALOG: &[FlutterField] = &[
    FlutterField::named("initialDate"),
    FlutterField::named("firstDate"),
    FlutterField::named("lastDate"),
    FlutterField::named("helpText"),
];

const F_TIMEPICKERDIALOG: &[FlutterField] = &[
    FlutterField::named("initialTime"),
    FlutterField::named("helpText"),
    FlutterField::named("cancelText"),
    FlutterField::named("confirmText"),
];

const F_TOOLTIP: &[FlutterField] = &[
    FlutterField::named("message"),
    FlutterField::named("child"),
    FlutterField::named("richMessage"),
    FlutterField::named("height"),
    FlutterField::named("padding"),
    FlutterField::named("waitDuration"),
    FlutterField::named("showDuration"),
];

const F_ICONBUTTON: &[FlutterField] = &[
    FlutterField::named("icon"),
    FlutterField::named("onPressed"),
    FlutterField::named("iconSize"),
    FlutterField::named("color"),
    FlutterField::named("tooltip"),
];

const F_OUTLINEDBTN: &[FlutterField] = &[
    FlutterField::named("onPressed"),
    FlutterField::named("child"),
    FlutterField::named("enabled"),
    FlutterField::named("style"),
    FlutterField::named("focusNode"),
    FlutterField::named("icon"),
    FlutterField::named("label"),
];

const F_TEXTBUTTON: &[FlutterField] = &[
    FlutterField::named("onPressed"),
    FlutterField::named("child"),
    FlutterField::named("enabled"),
    FlutterField::named("style"),
    FlutterField::named("autofocus"),
    FlutterField::named("icon"),
    FlutterField::named("label"),
];

const F_FAB: &[FlutterField] = &[
    FlutterField::named("onPressed"),
    FlutterField::named("child"),
    FlutterField::named("tooltip"),
    FlutterField::named("backgroundColor"),
    FlutterField::named("elevation"),
    FlutterField::named("foregroundColor"),
    FlutterField::named("mini"),
    FlutterField::named("shape"),
];

const F_POPUPMENUBTN: &[FlutterField] = &[
    FlutterField::named("itemBuilder"),
    FlutterField::named("initialValue"),
    FlutterField::named("onSelected"),
    FlutterField::named("icon"),
];

const F_POPUPMENUITEM: &[FlutterField] =
    &[FlutterField::named("value"), FlutterField::named("child")];

const F_BUTTONSTYLE: &[FlutterField] = &[
    FlutterField::named("backgroundColor"),
    FlutterField::named("elevation"),
    FlutterField::named("foregroundColor"),
];

const F_ROUNDEDBORDER: &[FlutterField] = &[
    FlutterField::named("borderRadius"),
    FlutterField::named("side"),
];

pub(crate) const CLASSES: &[FlutterClass] = &[
    // The app shell stacks its slots vertically: app bar, body, bottom bar.
    // `flex:1` because a `Scaffold` FILLS its app — Flutter hands it a tight
    // constraint, and without saying so its height was `auto`, so it shrank to
    // its content and the `Expanded` rows below had no height to divide.
    FlutterClass::widget(
        "Scaffold",
        "StatefulWidget",
        "div;display:flex;flex-direction:column;flex:1",
        SCAFFOLD_FIELDS,
    ),
    // An app bar IS the page header, and `<header>` already has a container
    // `control_kind` arm. `flex:0` keeps it a thin fixed bar while the body
    // takes the remaining height — the rule the old realizer applied by
    // special-casing the type name at realize time.
    // …and a bar is 56px of Material primary with a white title, which is what
    // makes it read as a bar rather than as a stray line of text above the
    // page. Nothing else declares this: `backgroundColor`/`elevation` arrive
    // null from a program that leaves them to the theme, and there is no theme.
    FlutterClass::widget(
        "AppBar",
        "StatefulWidget",
        "header;display:flex;align-items:center;flex:0;height:56px;\
         padding-left:16px;padding-right:16px;\
         background-color:#1976d2;color:#ffffff;font-size:20px;font-weight:bold",
        APPBAR_FIELDS,
    ),
    // **The app fills the window** — `height:100%`, the one declaration that
    // makes every constraint below it definite. Flutter's root has the window's
    // size by construction; a `<div>` does not, and with `height:auto` the
    // whole tree shrank to its content, which is a legal CSS reading of a
    // program that never said otherwise.
    FlutterClass::widget(
        "MaterialApp",
        "StatefulWidget",
        "div;display:flex;flex-direction:column;height:100%",
        MATERIALAPP_FIELDS,
    ),
    // **A Flutter button FILLS the box it is given** — that is what makes
    // `Expanded(child: Padding(child: ElevatedButton(…)))` a calculator key
    // rather than a label-sized button parked at the left of an empty cell.
    // A bare `<button>` is inline-block and shrinks to fit, which is right for
    // HTML and wrong for this widget, so the widget says so.
    FlutterClass::widget(
        "ElevatedButton",
        "StatefulWidget",
        "button;width:100%;height:100%",
        ELEVATEDBUTTON_FIELDS,
    ),
    // A card is a grouped surface — `<fieldset>` is what plib and dotnet both
    // map their group boxes to, so the three frontends stay one control.
    FlutterClass::widget("Card", "StatelessWidget", "fieldset", F_CARD),
    FlutterClass::widget("Drawer", "StatelessWidget", "aside", F_DRAWER),
    FlutterClass::widget("DrawerHeader", "StatelessWidget", "header", F_DRAWERHEADER),
    // A divider is a thematic break, which is what `<hr>` IS.
    FlutterClass::widget("Divider", "StatelessWidget", "hr", F_DIVIDER),
    FlutterClass::widget("VerticalDivider", "StatelessWidget", "hr", F_VDIVIDER),
    // A tile lays leading/title/subtitle/trailing out in a row.
    FlutterClass::widget(
        "ListTile",
        "StatelessWidget",
        "div;display:flex;flex-direction:row;align-items:center",
        F_LISTTILE,
    ),
    // Chips are labelled, tappable pills — a button both shows the label as
    // its face and delivers the tap.
    FlutterClass::widget("Chip", "StatelessWidget", "button", F_CHIP),
    FlutterClass::widget("ActionChip", "StatelessWidget", "button", F_ACTIONCHIP),
    // Filter/Choice chips are selectable — a checkbox carries the toggle state
    // a plain button cannot.
    FlutterClass::widget("FilterChip", "StatelessWidget", "input:checkbox", F_SELCHIP),
    FlutterClass::widget("ChoiceChip", "StatelessWidget", "input:checkbox", F_SELCHIP),
    // An avatar is an image, not a drawing surface.
    FlutterClass::widget("CircleAvatar", "StatelessWidget", "img", F_CIRCLEAVATAR),
    FlutterClass::widget("GridTile", "StatelessWidget", "div", F_GRIDTILE),
    FlutterClass::widget("GridTileBar", "StatelessWidget", "div", F_GRIDTILEBAR),
    FlutterClass::widget(
        "Stepper",
        "StatefulWidget",
        "div;display:flex;flex-direction:column",
        F_STEPPER,
    ),
    // HTML has tables outright, and `control_kind` maps `<table>` to the data
    // grid — the same control dotnet's `DataGridView` reaches.
    FlutterClass::widget("DataTable", "StatelessWidget", "table", F_DATATABLE),
    // A navigation bar IS `<nav>`, laid out as a row.
    FlutterClass::widget(
        "BottomNavigationBar",
        "StatefulWidget",
        "nav;display:flex;flex-direction:row",
        F_BOTTOMNAV,
    ),
    FlutterClass::widget("BottomAppBar", "StatefulWidget", "footer", F_BOTTOMAPPBAR),
    // A bottom sheet is a modal surface — `<dialog>` is the element with the
    // top-layer behaviour already landed.
    FlutterClass::widget("BottomSheet", "StatefulWidget", "dialog", F_BOTTOMSHEET),
    // ⚠ No HTML counterpart: a DECLARED custom element. `control_kind` strips
    // `vybe-` and looks the remainder up against the widget list, and
    // `tabcontrol` is a real kind — so the tag carries it.
    FlutterClass::widget("TabBar", "StatefulWidget", "vybe-tabcontrol", F_TABBAR),
    FlutterClass::widget("TabBarView", "StatefulWidget", "div", F_TABBARVIEW),
    // `Tab` contributes a LABEL to the tabcontrol's own header — it must not
    // realize a widget of its own, or each tab renders a second time on top of
    // the strip the tabcontrol already drew. Same idiom as `DropdownMenuItem`,
    // which maps to `Panel` so its text becomes an item rather than a control.
    FlutterClass::widget("Tab", "StatelessWidget", "vybe-tabpage", F_TAB),
    FlutterClass::widget("FlexibleSpaceBar", "StatefulWidget", "div", F_FLEXSPACEBAR),
    // HTML announces progress natively.
    FlutterClass::widget(
        "LinearProgressIndicator",
        "StatefulWidget",
        "progress",
        F_LINEARPROG,
    ),
    FlutterClass::widget(
        "CircularProgressIndicator",
        "StatefulWidget",
        "progress",
        F_CIRCPROG,
    ),
    // A picker dialog is a dialog holding a date input.
    FlutterClass::widget(
        "DatePickerDialog",
        "StatefulWidget",
        "input:date",
        F_DATEPICKERDIALOG,
    ),
    FlutterClass::widget(
        "TimePickerDialog",
        "StatefulWidget",
        "input:time",
        F_TIMEPICKERDIALOG,
    ),
    // A tooltip is a hover affordance around its child, not a box of its own.
    FlutterClass::wrapper("Tooltip", "StatefulWidget", F_TOOLTIP),
    FlutterClass::widget("BackButton", "StatelessWidget", "button", &[]),
    FlutterClass::widget("IconButton", "StatelessWidget", "button", F_ICONBUTTON),
    FlutterClass::widget("OutlinedButton", "StatefulWidget", "button", F_OUTLINEDBTN),
    FlutterClass::widget("TextButton", "StatefulWidget", "button", F_TEXTBUTTON),
    FlutterClass::widget("FloatingActionButton", "StatelessWidget", "button", F_FAB),
    FlutterClass::widget("PopupMenuButton", "StatefulWidget", "button", F_POPUPMENUBTN),
    // A menu ITEM is the same `<menu>` element as the strip that opens it —
    // the spelling plib and dotnet both use for `TMenuItem` /
    // `ToolStripMenuItem`.
    FlutterClass::widget("PopupMenuItem", "Widget", "menu", F_POPUPMENUITEM),
    FlutterClass::data("Step", None, F_STEP),
    FlutterClass::data("DataColumn", None, F_DATACOLUMN),
    FlutterClass::data("DataRow", None, F_DATAROW),
    FlutterClass::data("DataCell", None, F_DATACELL),
    FlutterClass::data("BottomNavigationBarItem", None, F_BOTTOMNAVITEM),
    FlutterClass::data("ButtonStyle", None, F_BUTTONSTYLE),
    FlutterClass::data("RoundedRectangleBorder", None, F_ROUNDEDBORDER),
    FlutterClass::data("ThemeData", None, &[]),
    FlutterClass::data("CircularNotchedRectangle", None, &[]),
    FlutterClass::data("HeroController", None, &[]),
];
