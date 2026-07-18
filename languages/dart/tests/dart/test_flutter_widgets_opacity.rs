use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Opacity
// ═══════════════════════════════════════════════════════════

#[test]
fn opacity_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final o = Opacity(opacity: 0.5, child: const SizedBox());
  print(o.opacity);
}
"#
        ),
        vec!["0.5"]
    );
}

#[test]
fn opacity_always_include_semantics() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final o = Opacity(opacity: 0.0, alwaysIncludeSemantics: true, child: const SizedBox());
  print(o.alwaysIncludeSemantics);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn opacity_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final o = Opacity(opacity: 1.0);
  print(o is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn animated_opacity_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ao = AnimatedOpacity(
    opacity: 0.5,
    duration: const Duration(milliseconds: 500),
    child: const SizedBox(),
  );
  print(ao.duration.inMilliseconds);
}
"#
        ),
        vec!["500"]
    );
}

#[test]
fn animated_opacity_curve() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ao = AnimatedOpacity(
    opacity: 0.5,
    duration: const Duration(seconds: 1),
    curve: Curves.easeOut,
  );
  print(ao.curve == Curves.easeOut);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_opacity_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final so = SliverOpacity(
    opacity: 0.8,
    sliver: const SliverToBoxAdapter(),
  );
  print(so.opacity);
}
"#
        ),
        vec!["0.8"]
    );
}

#[test]
fn sliver_opacity_always_include_semantics() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final so = SliverOpacity(
    opacity: 0.0,
    alwaysIncludeSemantics: true,
    sliver: const SliverToBoxAdapter(),
  );
  print(so.alwaysIncludeSemantics);
}
"#
        ),
        vec!["true"]
    );
}
