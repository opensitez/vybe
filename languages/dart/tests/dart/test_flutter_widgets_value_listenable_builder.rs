use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets ValueListenableBuilder
// ═══════════════════════════════════════════════════════════

#[test]
fn value_listenable_builder_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final vn = ValueNotifier<int>(0);
  final vlb = ValueListenableBuilder<int>(
    valueListenable: vn,
    builder: (BuildContext context, int value, Widget? child) {
      return const SizedBox();
    },
  );
  print(vlb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn value_listenable_builder_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final vn = ValueNotifier<String>('Test');
  final vlb = ValueListenableBuilder<String>(
    valueListenable: vn,
    child: const Text('Child'),
    builder: (BuildContext context, String value, Widget? child) {
      return child ?? const SizedBox();
    },
  );
  print(vlb.child is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn value_listenable_builder_listenable() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final vn = ValueNotifier<bool>(true);
  final vlb = ValueListenableBuilder<bool>(
    valueListenable: vn,
    builder: (BuildContext context, bool value, Widget? child) {
      return const SizedBox();
    },
  );
  print(vlb.valueListenable == vn);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn value_listenable_builder_builder() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final vn = ValueNotifier<double>(3.14);
  final vlb = ValueListenableBuilder<double>(
    valueListenable: vn,
    builder: (BuildContext context, double value, Widget? child) {
      return const SizedBox();
    },
  );
  print(vlb.builder != null);
}
"#
        ),
        vec!["true"]
    );
}
