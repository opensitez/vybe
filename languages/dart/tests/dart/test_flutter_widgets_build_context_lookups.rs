use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets BuildContext Lookups
// ═══════════════════════════════════════════════════════════

#[test]
fn context_find_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyWidget extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    // In real app, this finds the render object
    print(context.findRenderObject() == null);
    return const SizedBox();
  }
}
void main() {
  final w = MyWidget();
  final e = w.createElement();
  // e.findRenderObject() will return null because not mounted
  print(e.findRenderObject() == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn context_size() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = const SizedBox();
  final e = w.createElement();
  try {
    print(e.size);
  } catch(err) {
    // throws if not mounted or no render object
    print('throws');
  }
}
"#
        ),
        vec!["throws"]
    );
}

#[test]
fn context_depend_on_inherited_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyInherited extends InheritedWidget {
  MyInherited() : super(child: const SizedBox());
  @override
  bool updateShouldNotify(MyInherited old) => false;
}
void main() {
  final w = const SizedBox();
  final e = w.createElement();
  // returns null if not found
  final res = e.dependOnInheritedWidgetOfExactType<MyInherited>();
  print(res == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn context_get_element_for_inherited_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyInherited extends InheritedWidget {
  MyInherited() : super(child: const SizedBox());
  @override
  bool updateShouldNotify(MyInherited old) => false;
}
void main() {
  final w = const SizedBox();
  final e = w.createElement();
  final res = e.getElementForInheritedWidgetOfExactType<MyInherited>();
  print(res == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn context_find_ancestor_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = const SizedBox();
  final e = w.createElement();
  final res = e.findAncestorWidgetOfExactType<Placeholder>();
  print(res == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn context_find_ancestor_state() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = const SizedBox();
  final e = w.createElement();
  final res = e.findAncestorStateOfType<State<StatefulWidget>>();
  print(res == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn context_find_root_ancestor_state() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = const SizedBox();
  final e = w.createElement();
  final res = e.findRootAncestorStateOfType<State<StatefulWidget>>();
  print(res == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn context_find_ancestor_render_object_type() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final w = const SizedBox();
  final e = w.createElement();
  final res = e.findAncestorRenderObjectOfType<RenderBox>();
  print(res == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn context_visit_ancestor_elements() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = const SizedBox();
  final e = w.createElement();
  bool visited = false;
  e.visitAncestorElements((element) {
    visited = true;
    return false; // stop
  });
  print(visited); // No ancestors, should be false
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn context_visit_child_elements() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = const SizedBox();
  final e = w.createElement();
  bool visited = false;
  e.visitChildElements((element) {
    visited = true;
  });
  print(visited); // Not mounted, no children
}
"#
        ),
        vec!["false"]
    );
}
