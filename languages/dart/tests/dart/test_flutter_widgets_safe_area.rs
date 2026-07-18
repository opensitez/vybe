use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets SafeArea
// ═══════════════════════════════════════════════════════════

#[test]
fn safe_area_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sa = SafeArea(child: const SizedBox());
  print(sa != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn safe_area_left_padding() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sa = SafeArea(left: false, child: const SizedBox());
  print(sa.left);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn safe_area_top_padding() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sa = SafeArea(top: false, child: const SizedBox());
  print(sa.top);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn safe_area_right_padding() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sa = SafeArea(right: false, child: const SizedBox());
  print(sa.right);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn safe_area_bottom_padding() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sa = SafeArea(bottom: false, child: const SizedBox());
  print(sa.bottom);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn safe_area_minimum() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sa = SafeArea(minimum: EdgeInsets.all(10.0), child: const SizedBox());
  print(sa.minimum == EdgeInsets.all(10.0));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn safe_area_maintain_bottom_view_padding() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sa = SafeArea(maintainBottomViewPadding: true, child: const SizedBox());
  print(sa.maintainBottomViewPadding);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn safe_area_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sa = SafeArea(child: const Placeholder());
  print(sa.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn safe_area_is_stateless_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sa = SafeArea(child: const SizedBox());
  print(sa is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_safe_area_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ssa = SliverSafeArea(sliver: const SliverToBoxAdapter());
  print(ssa != null);
}
"#
        ),
        vec!["true"]
    );
}
