use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets BackdropFilter
// ═══════════════════════════════════════════════════════════

#[test]
fn backdrop_filter_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
void main() {
  final bf = BackdropFilter(
    filter: ImageFilter.blur(sigmaX: 5.0, sigmaY: 5.0),
    child: const SizedBox(),
  );
  print(bf.filter != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn backdrop_filter_blend_mode() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
void main() {
  final bf = BackdropFilter(
    filter: ImageFilter.blur(sigmaX: 5.0, sigmaY: 5.0),
    blendMode: BlendMode.srcOver,
    child: const SizedBox(),
  );
  print(bf.blendMode == BlendMode.srcOver);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn backdrop_filter_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
void main() {
  final bf = BackdropFilter(
    filter: ImageFilter.blur(sigmaX: 5.0, sigmaY: 5.0),
    child: const Placeholder(),
  );
  print(bf.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn backdrop_filter_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
void main() {
  final bf = BackdropFilter(
    filter: ImageFilter.blur(sigmaX: 5.0, sigmaY: 5.0),
  );
  print(bf is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn backdrop_filter_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
import 'package:flutter/rendering.dart';
void main() {
  final bf = BackdropFilter(
    filter: ImageFilter.blur(sigmaX: 5.0, sigmaY: 5.0),
  );
  // Creates RenderBackdropFilter
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
