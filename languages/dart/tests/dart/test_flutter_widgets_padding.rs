use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Padding
// ═══════════════════════════════════════════════════════════

#[test]
fn padding_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Padding(padding: EdgeInsets.all(8.0), child: const SizedBox());
  print(p.padding is EdgeInsets);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn padding_all() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Padding(padding: EdgeInsets.all(10.0));
  print('${p.padding.left}:${p.padding.right}');
}
"#
        ),
        vec!["10.0:10.0"]
    );
}

#[test]
fn padding_symmetric() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Padding(padding: EdgeInsets.symmetric(horizontal: 10.0, vertical: 20.0));
  print('${p.padding.left}:${p.padding.top}');
}
"#
        ),
        vec!["10.0:20.0"]
    );
}

#[test]
fn padding_only() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Padding(padding: EdgeInsets.only(left: 10.0, bottom: 20.0));
  print('${p.padding.left}:${p.padding.top}:${p.padding.right}:${p.padding.bottom}');
}
"#
        ),
        vec!["10.0:0.0:0.0:20.0"]
    );
}

#[test]
fn padding_is_single_child_render_object_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Padding(padding: EdgeInsets.zero);
  print(p is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn padding_child_retrieval() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Padding(padding: EdgeInsets.zero, child: const Placeholder());
  print(p.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn animated_padding_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ap = AnimatedPadding(
    padding: EdgeInsets.all(10),
    duration: const Duration(seconds: 1),
  );
  print(ap.duration.inSeconds);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn animated_padding_curve() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ap = AnimatedPadding(
    padding: EdgeInsets.all(10),
    duration: const Duration(seconds: 1),
    curve: Curves.easeIn,
  );
  print(ap.curve == Curves.easeIn);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_padding_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sp = SliverPadding(padding: EdgeInsets.all(8.0));
  print(sp.padding is EdgeInsets);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn edge_insets_geometry_abstract() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Padding(padding: EdgeInsets.all(10));
  print(p.padding is EdgeInsetsGeometry);
}
"#
        ),
        vec!["true"]
    );
}
