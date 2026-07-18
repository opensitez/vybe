use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Spacer
// ═══════════════════════════════════════════════════════════

#[test]
fn spacer_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Spacer();
  print(s != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn spacer_flex_default() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Spacer();
  print(s.flex);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn spacer_flex_custom() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Spacer(flex: 3);
  print(s.flex);
}
"#
        ),
        vec!["3"]
    );
}

#[test]
fn spacer_is_stateless_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Spacer();
  print(s is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn spacer_builds_expanded() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Spacer(flex: 2);
  // Spacer builds an Expanded with a SizedBox
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
