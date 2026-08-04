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

const F_ACTIONCHIP: &[FlutterField] =
    &[FlutterField::named("label"), FlutterField::named("onPressed")];

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

const F_DATAROW: &[FlutterField] =
    &[FlutterField::named("cells"), FlutterField::named("selected")];

const F_DATACELL: &[FlutterField] =
    &[FlutterField::positional("child", 0), FlutterField::named("showEditIcon")];

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

const F_ROUNDEDBORDER: &[FlutterField] =
    &[FlutterField::named("borderRadius"), FlutterField::named("side")];

pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::widget("Scaffold", "StatefulWidget", "FlowLayoutPanel", SCAFFOLD_FIELDS),
    FlutterClass::widget("AppBar", "StatefulWidget", "FlowLayoutPanel", APPBAR_FIELDS),
    FlutterClass::widget("MaterialApp", "StatefulWidget", "FlowLayoutPanel", MATERIALAPP_FIELDS),
    FlutterClass::widget("ElevatedButton", "StatefulWidget", "Button", ELEVATEDBUTTON_FIELDS),
    FlutterClass::widget("Card", "StatelessWidget", "groupbox", F_CARD),
    FlutterClass::widget("Drawer", "StatelessWidget", "Panel", F_DRAWER),
    FlutterClass::widget("DrawerHeader", "StatelessWidget", "Panel", F_DRAWERHEADER),
    FlutterClass::widget("Divider", "StatelessWidget", "Panel", F_DIVIDER),
    FlutterClass::widget("VerticalDivider", "StatelessWidget", "Panel", F_VDIVIDER),
    // A tile lays leading/title/subtitle/trailing out in a row, so it needs a
    // flow panel — a bare Panel gave its children no layout at all.
    FlutterClass::widget("ListTile", "StatelessWidget", "HFlowLayoutPanel", F_LISTTILE),
    // Chips are labelled, tappable pills — Button is the control that both
    // shows the label as its face and delivers the tap.
    FlutterClass::widget("Chip", "StatelessWidget", "Button", F_CHIP),
    FlutterClass::widget("ActionChip", "StatelessWidget", "Button", F_ACTIONCHIP),
    // Filter/Choice chips are selectable — a checkbox carries the toggle state
    // the plain Button cannot.
    FlutterClass::widget("FilterChip", "StatelessWidget", "CheckBox", F_SELCHIP),
    FlutterClass::widget("ChoiceChip", "StatelessWidget", "CheckBox", F_SELCHIP),
    FlutterClass::widget("CircleAvatar", "StatelessWidget", "picturebox", F_CIRCLEAVATAR),
    FlutterClass::widget("GridTile", "StatelessWidget", "Panel", F_GRIDTILE),
    FlutterClass::widget("GridTileBar", "StatelessWidget", "Panel", F_GRIDTILEBAR),
    FlutterClass::widget("Stepper", "StatefulWidget", "Panel", F_STEPPER),
    FlutterClass::widget("DataTable", "StatelessWidget", "datagrid", F_DATATABLE),
    FlutterClass::widget("BottomNavigationBar", "StatefulWidget", "HFlowLayoutPanel", F_BOTTOMNAV),
    FlutterClass::widget("BottomAppBar", "StatefulWidget", "Panel", F_BOTTOMAPPBAR),
    FlutterClass::widget("BottomSheet", "StatefulWidget", "Panel", F_BOTTOMSHEET),
    FlutterClass::widget("TabBar", "StatefulWidget", "tabcontrol", F_TABBAR),
    FlutterClass::widget("TabBarView", "StatefulWidget", "FlowLayoutPanel", F_TABBARVIEW),
    // `Tab` contributes a LABEL to the tabcontrol's own header — it must not
    // realize a widget of its own, or each tab renders a second time on top of
    // the strip the tabcontrol already drew. Same idiom as `DropdownMenuItem`,
    // which maps to `Panel` so its text becomes an item rather than a control.
    FlutterClass::widget("Tab", "StatelessWidget", "Panel", F_TAB),
    FlutterClass::widget("FlexibleSpaceBar", "StatefulWidget", "Panel", F_FLEXSPACEBAR),
    FlutterClass::widget("LinearProgressIndicator", "StatefulWidget", "progressbar", F_LINEARPROG),
    FlutterClass::widget("CircularProgressIndicator", "StatefulWidget", "progressbar", F_CIRCPROG),
    FlutterClass::widget("DatePickerDialog", "StatefulWidget", "datetimepicker", F_DATEPICKERDIALOG),
    FlutterClass::widget("TimePickerDialog", "StatefulWidget", "datetimepicker", F_TIMEPICKERDIALOG),
    // A tooltip is a hover affordance around its child, not a box of its own.
    FlutterClass::wrapper("Tooltip", "StatefulWidget", F_TOOLTIP),
    FlutterClass::widget("BackButton", "StatelessWidget", "Button", &[]),
    FlutterClass::widget("IconButton", "StatelessWidget", "Button", F_ICONBUTTON),
    FlutterClass::widget("OutlinedButton", "StatefulWidget", "Button", F_OUTLINEDBTN),
    FlutterClass::widget("TextButton", "StatefulWidget", "Button", F_TEXTBUTTON),
    FlutterClass::widget("FloatingActionButton", "StatelessWidget", "Button", F_FAB),
    FlutterClass::widget("PopupMenuButton", "StatefulWidget", "Button", F_POPUPMENUBTN),
    FlutterClass::widget("PopupMenuItem", "Widget", "Panel", F_POPUPMENUITEM),
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
