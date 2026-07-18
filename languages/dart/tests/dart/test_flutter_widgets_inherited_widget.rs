use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets InheritedWidget
// ═══════════════════════════════════════════════════════════

#[test]
fn inherited_widget_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyInherited extends InheritedWidget {
  final int value;
  MyInherited({required this.value, required Widget child}) : super(child: child);
  @override
  bool updateShouldNotify(MyInherited oldWidget) {
    return oldWidget.value != value;
  }
}
void main() {
  final w = MyInherited(value: 42, child: const SizedBox());
  print(w != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn inherited_widget_update_should_notify_true() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyInherited extends InheritedWidget {
  final int value;
  MyInherited({required this.value, required Widget child}) : super(child: child);
  @override
  bool updateShouldNotify(MyInherited oldWidget) {
    return oldWidget.value != value;
  }
}
void main() {
  final w1 = MyInherited(value: 1, child: const SizedBox());
  final w2 = MyInherited(value: 2, child: const SizedBox());
  print(w2.updateShouldNotify(w1));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn inherited_widget_update_should_notify_false() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyInherited extends InheritedWidget {
  final int value;
  MyInherited({required this.value, required Widget child}) : super(child: child);
  @override
  bool updateShouldNotify(MyInherited oldWidget) {
    return oldWidget.value != value;
  }
}
void main() {
  final w1 = MyInherited(value: 1, child: const SizedBox());
  final w2 = MyInherited(value: 1, child: const SizedBox());
  print(w2.updateShouldNotify(w1));
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn inherited_element_creation() {
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
  final w = MyInherited();
  final e = w.createElement();
  print(e is InheritedElement);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn proxy_widget_abstract() {
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
  final w = MyInherited();
  print(w is ProxyWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn inherited_notifier_subclass() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/foundation.dart';
class MyNotifier extends InheritedNotifier<ValueNotifier<int>> {
  MyNotifier({required ValueNotifier<int> notifier, required Widget child})
      : super(notifier: notifier, child: child);
}
void main() {
  final vn = ValueNotifier<int>(0);
  final w = MyNotifier(notifier: vn, child: const SizedBox());
  print(w != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn inherited_model_subclass() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyModel extends InheritedModel<String> {
  MyModel({required Widget child}) : super(child: child);
  @override
  bool updateShouldNotify(MyModel oldWidget) => true;
  @override
  bool updateShouldNotifyDependent(MyModel oldWidget, Set<String> dependencies) {
    return dependencies.contains('update');
  }
}
void main() {
  final w = MyModel(child: const SizedBox());
  print(w != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn inherited_model_update_should_notify_dependent() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyModel extends InheritedModel<String> {
  MyModel({required Widget child}) : super(child: child);
  @override
  bool updateShouldNotify(MyModel oldWidget) => true;
  @override
  bool updateShouldNotifyDependent(MyModel oldWidget, Set<String> dependencies) {
    return dependencies.contains('foo');
  }
}
void main() {
  final m1 = MyModel(child: const SizedBox());
  final m2 = MyModel(child: const SizedBox());
  print(m2.updateShouldNotifyDependent(m1, {'foo'}));
  print(m2.updateShouldNotifyDependent(m1, {'bar'}));
}
"#
        ),
        vec!["true\nfalse"]
    );
}

#[test]
fn of_method_pattern() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyInherited extends InheritedWidget {
  final int value = 42;
  MyInherited({required Widget child}) : super(child: child);
  @override
  bool updateShouldNotify(MyInherited old) => false;
  static MyInherited? of(BuildContext context) {
    return context.dependOnInheritedWidgetOfExactType<MyInherited>();
  }
}
void main() {
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
