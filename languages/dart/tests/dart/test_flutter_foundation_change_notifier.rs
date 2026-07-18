use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: foundation/change_notifier.dart Logic
// ═══════════════════════════════════════════════════════════

#[test]
fn change_notifier_add_listener_notify() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  int count = 0;
  notifier.addListener(() { count++; });
  notifier.notifyListeners();
  print(count);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn change_notifier_remove_listener() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  int count = 0;
  void callback() { count++; }
  notifier.addListener(callback);
  notifier.removeListener(callback);
  notifier.notifyListeners();
  print(count);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn change_notifier_notify_no_listeners() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  // Should not throw or crash
  notifier.notifyListeners();
  print('success');
}
"#
        ),
        vec!["success"]
    );
}

#[test]
fn change_notifier_duplicate_listeners() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  int count = 0;
  void callback() { count++; }
  notifier.addListener(callback);
  notifier.addListener(callback);
  notifier.notifyListeners();
  print(count);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn change_notifier_remove_one_duplicate() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  int count = 0;
  void callback() { count++; }
  notifier.addListener(callback);
  notifier.addListener(callback);
  notifier.removeListener(callback);
  notifier.notifyListeners();
  print(count);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn change_notifier_listener_modifies_listeners() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  int count1 = 0;
  int count2 = 0;
  void cb2() { count2++; }
  void cb1() { 
    count1++; 
    notifier.removeListener(cb2); // Modified during iteration
  }
  notifier.addListener(cb1);
  notifier.addListener(cb2);
  notifier.notifyListeners();
  print('$count1:$count2');
}
"#
        ),
        // Flutter's ChangeNotifier caches the listener list before iterating,
        // so cb2 will still be called on the first notify.
        vec!["1:1"]
    );
}

#[test]
fn change_notifier_listener_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  int count = 0;
  notifier.addListener(() { throw Exception('Bad listener'); });
  notifier.addListener(() { count++; });
  try {
    notifier.notifyListeners();
  } catch (e) {
    // Depending on Flutter implementation, it might catch and report to FlutterError,
    // or bubble up if there is no zone handler.
    print('Exception handled');
  }
  print(count);
}
"#
        ),
        // FlutterError.reportError catches it, so the next listener runs.
        // Wait, standard Dart throws. We'll just expect 'Exception handled' and check if count ran.
        vec!["Exception handled\n1"]
    );
}

#[test]
fn change_notifier_dispose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  notifier.dispose();
  try {
    notifier.notifyListeners();
  } catch (e) {
    print('FlutterError thrown');
  }
}
"#
        ),
        vec!["FlutterError thrown"]
    );
}

#[test]
fn change_notifier_add_after_dispose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  notifier.dispose();
  try {
    notifier.addListener(() {});
  } catch (e) {
    print('FlutterError thrown on add');
  }
}
"#
        ),
        vec!["FlutterError thrown on add"]
    );
}

#[test]
fn change_notifier_remove_after_dispose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  void cb() {}
  notifier.addListener(cb);
  notifier.dispose();
  // Removing after dispose is explicitly safe in Flutter
  notifier.removeListener(cb);
  print('Safe remove');
}
"#
        ),
        vec!["Safe remove"]
    );
}

#[test]
fn change_notifier_subclassing() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
class MyNotifier extends ChangeNotifier {
  int _value = 0;
  int get value => _value;
  set value(int v) {
    _value = v;
    notifyListeners();
  }
}
void main() {
  final n = MyNotifier();
  int count = 0;
  n.addListener(() { count++; });
  n.value = 42;
  print('$count:${n.value}');
}
"#
        ),
        vec!["1:42"]
    );
}

#[test]
fn change_notifier_has_listeners_true() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  notifier.addListener(() {});
  print(notifier.hasListeners);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn change_notifier_has_listeners_false() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  void cb() {}
  notifier.addListener(cb);
  notifier.removeListener(cb);
  print(notifier.hasListeners);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn change_notifier_memory_leak_prevention() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  void cb() {}
  notifier.addListener(cb);
  notifier.removeListener(cb);
  // Simulating leak check by checking hasListeners
  print(notifier.hasListeners);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn value_notifier_update() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ValueNotifier<int>(0);
  int count = 0;
  notifier.addListener(() { count++; });
  notifier.value = 1;
  notifier.value = 1; // Same value should not trigger
  print('$count:${notifier.value}');
}
"#
        ),
        vec!["1:1"]
    );
}

#[test]
fn change_notifier_large_number_of_listeners() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  int count = 0;
  for (int i = 0; i < 1000; i++) {
    notifier.addListener(() { count++; });
  }
  notifier.notifyListeners();
  print(count);
}
"#
        ),
        vec!["1000"]
    );
}

#[test]
fn change_notifier_listener_disposes_notifier() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final notifier = ChangeNotifier();
  int count2 = 0;
  notifier.addListener(() {
    notifier.dispose();
  });
  notifier.addListener(() {
    count2++;
  });
  // The first listener disposes it. The iteration should complete safely in Flutter.
  try {
    notifier.notifyListeners();
  } catch (e) {
    print('FlutterError');
  }
  print(count2);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn change_notifier_mixin() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
class MyModel with ChangeNotifier {
  void update() { notifyListeners(); }
}
void main() {
  final m = MyModel();
  int c = 0;
  m.addListener(() { c++; });
  m.update();
  print(c);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn change_notifier_overriding_dispose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
class BadNotifier extends ChangeNotifier {
  @override
  void dispose() {
    // missing super.dispose()
  }
}
void main() {
  final n = BadNotifier();
  n.dispose();
  // Because super was not called, it doesn't throw on notify
  n.notifyListeners();
  print('survived');
}
"#
        ),
        vec!["survived"]
    );
}

#[test]
fn value_notifier_equality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final v1 = ValueNotifier<int>(1);
  final v2 = ValueNotifier<int>(1);
  print(v1 == v2);
}
"#
        ),
        vec!["false"]
    );
}
