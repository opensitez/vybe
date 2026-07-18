use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets AnimatedBuilder
// ═══════════════════════════════════════════════════════════

#[test]
fn animated_builder_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final vn = ValueNotifier<double>(0.0);
  final ab = AnimatedBuilder(
    animation: vn,
    builder: (BuildContext context, Widget? child) {
      return const SizedBox();
    },
  );
  print(ab is AnimatedWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn animated_builder_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final vn = ValueNotifier<double>(0.0);
  final ab = AnimatedBuilder(
    animation: vn,
    child: const Text('Child'),
    builder: (BuildContext context, Widget? child) {
      return child ?? const SizedBox();
    },
  );
  print(ab.child is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn animated_builder_animation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final vn = ValueNotifier<double>(0.0);
  final ab = AnimatedBuilder(
    animation: vn,
    builder: (BuildContext context, Widget? child) {
      return const SizedBox();
    },
  );
  print(ab.listenable == vn);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn animated_builder_builder_function() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final vn = ValueNotifier<double>(0.0);
  final ab = AnimatedBuilder(
    animation: vn,
    builder: (BuildContext context, Widget? child) {
      return const Text('Builder');
    },
  );
  print(ab.builder != null);
}
"#
        ),
        vec!["true"]
    );
}
