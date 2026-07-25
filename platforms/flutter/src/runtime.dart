// ─────────────────────────────────────────────────────────────────────────
// Vybe Flutter runtime adapter — injected only when a program renders (imports
// `package:flutter/*` AND references `runApp`).
//
// Every widget is a config OBJECT stamped by the compiler with `__controlfn`
// (its vybe:gui control type) and `__ops` (a [kind,key,value] instruction
// list). The realizer creates each control ONCE under a stable, tree-path name
// and keeps it in vybe_widgets; `setState` re-walks the same tree and just
// pushes changed property values onto the SAME named controls — no rebuild.
// The control's NAME is its identity (the same model WinForms/VCL use), so all
// adapters share one state store keyed by name.

dynamic _vfRoot;
dynamic _vfForm;
var _vfStates = {}; // State instances, keyed by widget type (persist rebuilds)

// A real widget (vs a scalar / value-type like EdgeInsets): a catalog widget
// carries `__controlfn`; a user composite carries `build`/`createState`.
bool _vfIsWidget(dynamic v) {
  if (v == null) {
    return false;
  }
  return v.__controlfn != null || v.build != null || v.createState != null;
}

// Expand a user composite (StatelessWidget/StatefulWidget) to its concrete
// catalog widget by running build(); State persists per widget type.
dynamic _vfConcrete(dynamic w) {
  while (w != null && w.__controlfn == null) {
    if (w.createState != null) {
      var st = _vfStates[w.__type];
      if (st == null) {
        st = w.createState();
        _vfStates[w.__type] = st;
      }
      w = st.build(null);
    } else {
      w = w.build(null);
    }
  }
  return w;
}

String _vfCaption(dynamic w) {
  if (w == null) {
    return "";
  }
  return "${w.data}";
}

// A control whose children are LIST ITEMS (combobox/listbox/tabcontrol/
// datagrid), not nested sub-controls. Their items are populated by caption
// via clearitems/additem rather than by realizing each child as a control.
bool _vfIsItemControl(dynamic cfn) {
  return cfn == "combobox" ||
      cfn == "listbox" ||
      cfn == "tabcontrol" ||
      cfn == "datagrid";
}

// The display caption for one list item — a DropdownMenuItem(child: Text),
// a ListTile(title: Text), a Tab(text:), or a bare Text.
String _vfItemCaption(dynamic item) {
  if (item == null) {
    return "";
  }
  if (item.child != null && item.child.data != null) {
    return "${item.child.data}";
  }
  if (item.title != null && item.title.data != null) {
    return "${item.title.data}";
  }
  if (item.text != null) {
    return "${item.text}";
  }
  if (item.data != null) {
    return "${item.data}";
  }
  return "${item}";
}

// Index of the item whose `.value` equals `val` (DropdownButton's selected
// value → the combobox's selected index). -1 if none / no items.
int _vfItemIndexOfValue(dynamic items, dynamic val) {
  if (items == null) {
    return -1;
  }
  for (var i = 0; i < items.length; i++) {
    if (items[i].value == val) {
      return i;
    }
  }
  return -1;
}

// Create-or-update the control for `w` at `path`, nesting into `parent` (a
// control, or null = the form root) only on first creation.
void _vfRealize(dynamic w, String path, dynamic parent) {
  w = _vfConcrete(w);
  if (w == null || w.__controlfn == null) {
    return;
  }
  var isNew = !vybe.gui.hasControl(path);
  var control = null;
  if (isNew) {
    control = vybe.gui.newControl(w.__controlfn, path);
    if (parent == null) {
      vybe.gui.controlsAdd(_vfForm, control);
    } else {
      vybe.gui.controlsAdd(parent, control);
    }
    // A Scaffold app-bar is a thin fixed bar; everything else flexes to fill.
    if (w.__type == "AppBar") {
      vybe.gui.setProperty(path, "flex", "0");
    }
  }
  var ops = w.__ops;
  for (var i = 0; i < ops.length; i++) {
    var op = ops[i];
    var kind = op[0];
    var key = op[1];
    var value = op[2];
    if (kind == 0) {
      // NestOrProp: a child widget nests; a scalar sets the property.
      if (_vfIsWidget(value)) {
        _vfRealize(value, path + "/" + key, control);
      } else if (value != null) {
        if (key == "controller") {
          // A `TextEditingController` carries the field's text — forward its
          // `.text` to the control, not the controller object. Other
          // controllers (Scroll/Page) have no `.text` and are skipped.
          if (value.text != null) {
            vybe.gui.setProperty(path, "text", value.text);
          }
        } else if (_vfIsItemControl(w.__controlfn) && key == "value") {
          // A DropdownButton's selected value → the combobox's selected index
          // (the item whose `.value` matches). Items are populated first
          // (the `items` field precedes `value` in field order).
          var idx = _vfItemIndexOfValue(w.items, value);
          if (idx >= 0) {
            vybe.gui.setProperty(path, "selectedindex", "$idx");
          }
        } else {
          vybe.gui.setProperty(path, key, value);
        }
      }
    } else if (kind == 1) {
      // Children list.
      if (value != null) {
        if (_vfIsItemControl(w.__controlfn)) {
          // Item-bearing control: children are ENTRIES, not sub-controls —
          // rebuild the item list by caption (clear so setState re-sync is
          // idempotent).
          vybe.gui.setProperty(path, "clearitems", "");
          for (var j = 0; j < value.length; j++) {
            vybe.gui.setProperty(path, "additem", _vfItemCaption(value[j]));
          }
        } else {
          for (var j = 0; j < value.length; j++) {
            _vfRealize(value[j], path + "/" + key + "$j", control);
          }
        }
      }
    } else if (kind == 2) {
      // Event handler (wire once, on creation).
      if (isNew && value != null) {
        vybe.gui.onEvent(path, key, value);
      }
    } else if (kind == 3) {
      // Caption: a child Text's data is this control's text. A captioned leaf
      // (button/app-bar) emits a caption op for BOTH `child` and `label`; skip
      // the absent one (null) so it can't clobber the caption that IS set.
      if (value != null) {
        vybe.gui.setProperty(path, "Text", _vfCaption(value));
      }
    }
  }
  // Radio: the control is SELECTED when its own `value` equals the group's
  // `groupValue`. Only radios carry BOTH fields, so this is a safe probe.
  if (w.groupValue != null && w.value != null) {
    vybe.gui.setProperty(path, "checked", (w.value == w.groupValue) ? "true" : "false");
  }
}

void _vfRealizeRoot() {
  _vfRealize(_vfRoot, "r", null);
}

void runApp(dynamic app) {
  _vfRoot = app;
  _vfForm = vybe.gui.createForm("App");
  vybe.gui.setProperty(_vfForm, "Width", 360);
  vybe.gui.setProperty(_vfForm, "Height", 560);
  _vfRealizeRoot();
  vybe.gui.runApplication(_vfForm);
}

// Flutter's `State.setState`: run the mutation, then re-walk the tree pushing
// changed values onto the existing (same-named) controls — no rebuild.
void setState(dynamic fn) {
  fn();
  if (_vfForm != null) {
    _vfRealizeRoot();
  }
}

// Minimal layout value types the samples reference.
class EdgeInsets {
  double left = 0.0;
  double top = 0.0;
  double right = 0.0;
  double bottom = 0.0;
  EdgeInsets.all(double v) {
    left = v;
    top = v;
    right = v;
    bottom = v;
  }
  EdgeInsets.symmetric(double horizontal, double vertical) {
    left = horizontal;
    right = horizontal;
    top = vertical;
    bottom = vertical;
  }
  EdgeInsets.only(double l, double t, double r, double b) {
    left = l;
    top = t;
    right = r;
    bottom = b;
  }
}

class Alignment {
  double x = 0.0;
  double y = 0.0;
  Alignment(double ax, double ay) {
    x = ax;
    y = ay;
  }
}
