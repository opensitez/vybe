//! Input widgets — text fields, selection controls (checkbox/radio/switch),
//! sliders, dropdowns, and their decoration/controller data types.

use crate::emitter::catalog::{FlutterClass, FlutterField};

const F_TEXTFIELD: &[FlutterField] = &[
    FlutterField::named("controller"),
    FlutterField::named("focusNode"),
    FlutterField::named("decoration"),
    FlutterField::named("keyboardType"),
    FlutterField::named("obscureText"),
    FlutterField::named("maxLines"),
    FlutterField::named("onChanged"),
];

const F_TEXTFORMFIELD: &[FlutterField] = &[
    FlutterField::named("controller"),
    FlutterField::named("initialValue"),
    FlutterField::named("validator"),
    FlutterField::named("onSaved"),
    FlutterField::named("decoration"),
    FlutterField::named("obscureText"),
];

const F_FORM: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("autovalidateMode"),
    FlutterField::named("onChanged"),
    FlutterField::named("canPop"),
    FlutterField::named("onPopInvokedWithResult"),
];

const F_CHECKBOX: &[FlutterField] = &[
    FlutterField::named("value"),
    FlutterField::named("onChanged"),
    FlutterField::named("tristate"),
    FlutterField::named("activeColor"),
    FlutterField::named("checkColor"),
    FlutterField::named("isError"),
    FlutterField::named("focusNode"),
];

const RADIO_FIELDS: &[FlutterField] = &[
    FlutterField::named("value"),
    FlutterField::named("groupValue"),
    FlutterField::named("onChanged"),
    FlutterField::named("activeColor"),
    FlutterField::named("toggleable"),
    FlutterField::named("focusNode"),
];

const F_SWITCH: &[FlutterField] = &[
    FlutterField::named("value"),
    FlutterField::named("onChanged"),
    FlutterField::named("activeColor"),
    FlutterField::named("activeTrackColor"),
    FlutterField::named("inactiveThumbColor"),
    FlutterField::named("inactiveTrackColor"),
];

const F_SLIDER: &[FlutterField] = &[
    FlutterField::named("value"),
    FlutterField::named("onChanged"),
    FlutterField::named("min"),
    FlutterField::named("max"),
    FlutterField::named("divisions"),
    FlutterField::named("activeColor"),
    FlutterField::named("inactiveColor"),
];

const F_DROPDOWNBTN: &[FlutterField] = &[
    FlutterField::children_list("items"),
    FlutterField::named("onChanged"),
    FlutterField::named("value"),
    FlutterField::named("icon"),
    FlutterField::named("isExpanded"),
];

const F_DROPDOWNITEM: &[FlutterField] =
    &[FlutterField::named("value"), FlutterField::named("child")];

const F_INPUTDECORATOR: &[FlutterField] = &[
    FlutterField::named("decoration"),
    FlutterField::named("child"),
    FlutterField::named("baseStyle"),
    FlutterField::named("isFocused"),
    FlutterField::named("isHovering"),
    FlutterField::named("expands"),
];

const F_INPUTDECORATION: &[FlutterField] = &[
    FlutterField::named("labelText"),
    FlutterField::named("hintText"),
    FlutterField::named("border"),
    FlutterField::named("icon"),
];

const F_TEXTEDITCTRL: &[FlutterField] = &[FlutterField::named("text")];

pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::widget("TextField", "StatefulWidget", "TextBox", F_TEXTFIELD),
    FlutterClass::widget("TextFormField", "FormField", "TextBox", F_TEXTFORMFIELD),
    FlutterClass::widget("Form", "StatefulWidget", "Panel", F_FORM),
    FlutterClass::widget("Checkbox", "StatefulWidget", "CheckBox", F_CHECKBOX),
    FlutterClass::widget("Radio", "StatefulWidget", "RadioButton", RADIO_FIELDS),
    FlutterClass::widget("Switch", "StatefulWidget", "CheckBox", F_SWITCH),
    FlutterClass::widget("Slider", "StatefulWidget", "trackbar", F_SLIDER),
    FlutterClass::widget("DropdownButton", "StatefulWidget", "combobox", F_DROPDOWNBTN),
    FlutterClass::widget("DropdownMenuItem", "Widget", "Panel", F_DROPDOWNITEM),
    FlutterClass::widget("InputDecorator", "StatefulWidget", "Panel", F_INPUTDECORATOR),
    FlutterClass::data("InputDecoration", None, F_INPUTDECORATION),
    FlutterClass::data("TextEditingController", None, F_TEXTEDITCTRL),
];
