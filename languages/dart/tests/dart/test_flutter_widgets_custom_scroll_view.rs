use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets CustomScrollView
// ═══════════════════════════════════════════════════════════

#[test]
fn custom_scroll_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final csv = CustomScrollView(slivers: const []);
  print(csv is ScrollView);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_scroll_view_scroll_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final csv = CustomScrollView(
    scrollDirection: Axis.horizontal,
    slivers: const [],
  );
  print(csv.scrollDirection == Axis.horizontal);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_scroll_view_reverse() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final csv = CustomScrollView(
    reverse: true,
    slivers: const [],
  );
  print(csv.reverse);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_scroll_view_primary() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final csv = CustomScrollView(
    primary: true,
    slivers: const [],
  );
  print(csv.primary);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_scroll_view_physics() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final csv = CustomScrollView(
    physics: const BouncingScrollPhysics(),
    slivers: const [],
  );
  print(csv.physics is BouncingScrollPhysics);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_scroll_view_anchor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final csv = CustomScrollView(
    anchor: 0.5,
    slivers: const [],
  );
  print(csv.anchor);
}
"#
        ),
        vec!["0.5"]
    );
}

#[test]
fn custom_scroll_view_center() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final Key centerKey = const ValueKey('center');
  final csv = CustomScrollView(
    center: centerKey,
    slivers: const [],
  );
  print((csv.center as ValueKey).value);
}
"#
        ),
        vec!["center"]
    );
}
