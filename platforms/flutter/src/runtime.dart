// ─────────────────────────────────────────────────────────────────────────
// Vybe Flutter runtime adapter — injected only when a program renders (imports
// `package:flutter/*` AND references `runApp`).
//
// A widget IS its element. The compiler builds it at the CONSTRUCTION site,
// where a field's declared role is still known (`primitives/expressions.rs`,
// `emit_gui_field`), so by the time this file runs the tree already exists,
// nested and configured. What is left for the runtime is small: inflate user
// composites (`build`/`createState`, which only the guest can run), attach the
// root to the document, and honour `setState`'s contract.
//
// ⛔ THE DOM IS THE AUTHORITY AND `web:*` IS THE WHOLE SURFACE. This adapter
// reaches no other host. Nothing here addresses a control by a tree-path NAME
// the way WinForms and VCL do: a widget IS its element, so there is nothing to
// look up.

dynamic _vfRoot;
var _vfStates = {}; // State instances, keyed by widget type (persist rebuilds)

// Expand a user composite (StatelessWidget/StatefulWidget) to its concrete
// catalog widget by running build(); State persists per widget type.
//
// ⚠ The test is POSITIVE — "is this still a composite" — not "does it lack a
// control fn". The negative spelling over-ran by one: a widget IS an element
// now, and `__controlfn` on an element is read late-bound, through the element
// attribute path, where it answers null. So the loop called `build` on a
// `MaterialApp` that has none, got back `undefined`, and `runApp` appended
// nothing — a blank form, no error. Asking what a thing HAS cannot fail that
// way.
//
// △ `_vfStates` is keyed by `w.__type`, so two instances of the SAME
// StatefulWidget class share one State object. That is wrong — Flutter gives
// every element its own State — but the key has no instance identity to use
// yet, and fixing it is a change to the element/State model rather than to
// this loop. Left as-is deliberately, and recorded, not hidden.
dynamic _vfConcrete(dynamic w) {
  while (w != null && (w.createState != null || w.build != null)) {
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

// `runApp` attaches the root widget to the document, and that is all it does.
//
// Constructing a widget ALREADY builds its element and applies its declared
// arguments, so by the time `runApp` is called the whole tree exists, nested
// and configured; what it lacks is a parent. Appending the root is the only
// step left.
//
// There is no form to create and no application to run: a page is not told to
// start. `createForm`/`runApplication` were the old host's shape and have no
// web meaning — a browser has no "run this app". Asking for the document here
// is the deliberate, EXPLICIT act that opens the browsing context.
void runApp(dynamic app) {
  _vfRoot = app;
  var doc = web.html.activeDocument();
  var root = _vfConcrete(app);
  if (root != null) {
    web.dom.appendChild(doc, web.html.body(doc), root);
  }
}

// Flutter's `State.setState`: run the mutation, then REBUILD.
//
// Both halves are needed, and the second used to be missing. The old body was
// just `fn()`, on the reasoning that a widget IS its element so a mutation has
// already reached the document. That holds for a mutation that writes to a
// widget and not for the ordinary case: `current` is a plain `String` field and
// the display is `Text(current)`, built ONCE at construction with the old
// value. Nothing re-ran `build`, so a calculator key press changed the state
// and redrew nothing — clicks worked and the screen never moved.
//
// Rebuilding means running `build` again, which constructs a fresh element
// tree (the compiler builds a widget's element at its construction site), and
// swapping it for the one in the body. `_vfStates` is keyed by widget type and
// so survives, which is what carries the mutated state into the new tree.
//
// ⚠ This replaces the WHOLE tree on every `setState` — new elements, new
// listeners, and the old subtree discarded. Flutter reconciles instead,
// keeping elements whose type and key match and updating only what differs.
// That needs an identity per element to diff against, the same identity
// `_vfStates` lacks (see `_vfConcrete`). Correct before cheap, and recorded
// rather than hidden: this is O(tree) per interaction.
void setState(dynamic fn) {
  fn();
  _vfRebuild();
}

// Re-run the root's `build` and RECONCILE the result into the document.
//
// The mounted tree is read from the body rather than remembered in a variable:
// the document is the authority on what is displayed, and a remembered handle
// is a second copy of that fact which can be wrong. It was — tracking it in
// `_vfMounted` left the old tree in place and appended the new one beside it,
// so a key press produced two calculators.
void _vfRebuild() {
  if (_vfRoot == null) {
    return;
  }
  var doc = web.html.activeDocument();
  var fresh = _vfConcrete(_vfRoot);
  if (fresh == null) {
    return;
  }
  var body = web.html.body(doc);
  var mounted = web.dom.firstElementChild(doc, body);
  if (mounted == null) {
    web.dom.appendChild(doc, body, fresh);
  } else {
    _vfPatch(doc, body, mounted, fresh);
  }
}

// Make `old` match `fresh`, in place — Flutter's reconciliation, over the DOM.
//
// Replacing the whole tree is correct and wasteful: it discards every element
// and every listener on each interaction, and a page that rebuilds on a
// keystroke rebuilds the world. Patching keeps the elements that are still
// right and touches only what changed, which is what the DOM's mutation
// methods are FOR.
//
// The match test is the node's name — Flutter compares runtime type and key,
// and the element's tag is the same question here, since a widget IS its
// element. Different name ⟹ a different box ⟹ swap it whole. Same name ⟹
// recurse.
//
// Listeners are deliberately NOT re-registered: the surviving element keeps the
// handler it was given, and that handler reads the State object, which
// `_vfStates` persists across rebuilds. Re-adding would double-fire.
void _vfPatch(dynamic doc, dynamic parent, dynamic old, dynamic fresh) {
  if (web.dom.nodeName(doc, old) != web.dom.nodeName(doc, fresh)) {
    web.dom.replaceChild(doc, parent, fresh, old);
    return;
  }
  _vfPatchAttributes(doc, old, fresh);
  // Snapshots, not live collections — the loops below MUTATE both trees, and a
  // live list would shift underneath the index doing the mutating.
  var oldKids = web.dom.children(doc, old);
  var newKids = web.dom.children(doc, fresh);
  var i = 0;
  while (i < oldKids.length && i < newKids.length) {
    _vfPatch(doc, old, oldKids[i], newKids[i]);
    i = i + 1;
  }
  // The new tree is longer: adopt the extra children.
  while (i < newKids.length) {
    web.dom.appendChild(doc, old, newKids[i]);
    i = i + 1;
  }
  // …or shorter: drop the tail, from the end, so earlier indices stay valid.
  var j = oldKids.length - 1;
  while (j >= newKids.length) {
    web.dom.removeChild(doc, old, oldKids[j]);
    j = j - 1;
  }
  // A leaf's text is the thing a rebuild actually moves — the calculator's
  // display, a cell's mark, a status line.
  if (newKids.length == 0) {
    var text = web.dom.textContent(doc, fresh);
    if (web.dom.textContent(doc, old) != text) {
      web.dom.setTextContent(doc, old, text);
    }
  }
}

// The attribute half of the diff — both directions.
//
// Set what the new tree has and the old one does not, or has differently; then
// remove what the old one still carries and the new one dropped. One direction
// alone is a leak: an attribute a rebuild REMOVED would stay on the element for
// ever, which is how a disabled button never re-enables.
//
// `getAttributeNames()` is what makes this possible at all — every other
// attribute call is addressed by name, so before it there was no way to ask an
// element what it had.
void _vfPatchAttributes(dynamic doc, dynamic old, dynamic fresh) {
  var wanted = web.dom.getAttributeNames(doc, fresh);
  var i = 0;
  while (i < wanted.length) {
    var name = wanted[i];
    var value = web.dom.getAttribute(doc, fresh, name);
    if (web.dom.getAttribute(doc, old, name) != value) {
      web.dom.setAttribute(doc, old, name, value);
    }
    i = i + 1;
  }
  var had = web.dom.getAttributeNames(doc, old);
  var j = 0;
  while (j < had.length) {
    var name = had[j];
    if (web.dom.getAttribute(doc, fresh, name) == null) {
      web.dom.removeAttribute(doc, old, name);
    }
    j = j + 1;
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

  // **What this becomes when it reaches CSS.** `padding` is a LENGTH role, so
  // the emitter appends `px` to whatever this stringifies to — which is why an
  // `EdgeInsets` with no `toString` arrived as the literal `[object]` and every
  // Flutter padding in the document did nothing at all.
  //
  // A uniform inset — `EdgeInsets.all(24)`, the overwhelmingly common case —
  // is one number, and `24` + `px` is the value CSS wants.
  //
  // ⚠ A NON-uniform inset needs the four-value shorthand (`top right bottom
  // left`), and that cannot survive a `px` appended to the end of it. Rather
  // than emit a single side and quietly get three of them wrong, this answers
  // nothing: `"" + "px"` is not a length, so the declaration is dropped and the
  // box keeps its default. Visibly unstyled beats silently mis-styled, and
  // closing it properly means the length roles learning that a value can arrive
  // WITH its units already on it.
  String toString() {
    if (left == top && top == right && right == bottom) {
      return "$left";
    }
    return "";
  }
}

class Alignment {
  double x = 0.0;
  double y = 0.0;
  Alignment(double ax, double ay) {
    x = ax;
    y = ay;
  }

  // **What this becomes in CSS**, for the case a `Container` uses it: where its
  // content sits across the box, which is `text-align`. VCL's `Alignment` and
  // WinForms' `TextAlign` already map to the same property, so this is the
  // spelling those frontends share rather than a Flutter-specific one.
  //
  // ⚠ The INLINE axis only. Flutter's `Alignment` is two axes, and `y` has no
  // counterpart here — `text-align` says nothing about the block direction. It
  // also aligns inline CONTENT, so `Container(alignment: …, child: Text(…))`
  // lands correctly and a container holding a BOX child does not move. Both
  // need the flex mapping (`justify-content` + `align-items`), which is a role
  // that does not exist yet.
  String toString() {
    if (x < -0.5) {
      return "left";
    }
    if (x > 0.5) {
      return "right";
    }
    return "center";
  }

  // Flutter's nine named alignments, on its own -1..1 axes. Without these
  // `Alignment.centerRight` resolved to NOTHING — the field arrived null and
  // the element carried a valueless `alignment` attribute, which is why a
  // right-aligned calculator display sat on the left.
  static Alignment get topLeft { return Alignment(-1.0, -1.0); }
  static Alignment get topCenter { return Alignment(0.0, -1.0); }
  static Alignment get topRight { return Alignment(1.0, -1.0); }
  static Alignment get centerLeft { return Alignment(-1.0, 0.0); }
  static Alignment get center { return Alignment(0.0, 0.0); }
  static Alignment get centerRight { return Alignment(1.0, 0.0); }
  static Alignment get bottomLeft { return Alignment(-1.0, 1.0); }
  static Alignment get bottomCenter { return Alignment(0.0, 1.0); }
  static Alignment get bottomRight { return Alignment(1.0, 1.0); }

}
