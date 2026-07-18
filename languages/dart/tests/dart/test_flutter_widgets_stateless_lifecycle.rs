use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Stateless Lifecycle
// ═══════════════════════════════════════════════════════════

#[test]
fn stateless_widget_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyWidget extends StatelessWidget {
  @override
  Widget build(BuildContext context) => Placeholder();
}
void main() {
  final w = MyWidget();
  print(w != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn stateless_widget_build_method() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyWidget extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    print('building');
    return const SizedBox();
  }
}
void main() {
  final w = MyWidget();
  // BuildContext mock is often just passing null in naive tests, though flutter will complain
  // We just test if method exists and is callable
  try {
    w.build(null as dynamic);
  } catch(e) {
    // some widgets assert context != null
    print('failed');
  }
}
"#
        ),
        vec!["building"]
    );
}

#[test]
fn stateless_element_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyWidget extends StatelessWidget {
  @override
  Widget build(BuildContext context) => const SizedBox();
}
void main() {
  final w = MyWidget();
  final e = w.createElement();
  print(e is StatelessElement);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn element_widget_property() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyWidget extends StatelessWidget {
  @override
  Widget build(BuildContext context) => const SizedBox();
}
void main() {
  final w = MyWidget();
  final e = w.createElement();
  print(e.widget == w);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn stateless_element_build() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyWidget extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return const Placeholder();
  }
}
void main() {
  final w = MyWidget();
  final e = w.createElement();
  // Normally flutter calls e.build() during mount
  final child = e.build();
  print(child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn widget_can_update() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w1 = const SizedBox();
  final w2 = const SizedBox();
  print(Widget.canUpdate(w1, w2));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn widget_can_update_different_types() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w1 = const SizedBox();
  final w2 = const Placeholder();
  print(Widget.canUpdate(w1, w2));
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn widget_can_update_different_keys() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w1 = SizedBox(key: ValueKey(1));
  final w2 = SizedBox(key: ValueKey(2));
  print(Widget.canUpdate(w1, w2));
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn widget_can_update_same_key() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final k = UniqueKey();
  final w1 = SizedBox(key: k);
  final w2 = SizedBox(key: k);
  print(Widget.canUpdate(w1, w2));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn widget_hashcode() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w1 = const SizedBox();
  final w2 = const SizedBox();
  print(w1.hashCode == w2.hashCode);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn build_context_interface() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  // Element implements BuildContext
  final w = const SizedBox();
  final e = w.createElement();
  print(e is BuildContext);
}
"#
        ),
        vec!["true"]
    );
}
