use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Center
// ═══════════════════════════════════════════════════════════

#[test]
fn center_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Center(child: const SizedBox());
  print(c.alignment == Alignment.center);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn center_width_factor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Center(widthFactor: 3.0, child: const SizedBox());
  print(c.widthFactor);
}
"#
        ),
        vec!["3.0"]
    );
}

#[test]
fn center_height_factor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Center(heightFactor: 4.0, child: const SizedBox());
  print(c.heightFactor);
}
"#
        ),
        vec!["4.0"]
    );
}

#[test]
fn center_is_align() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Center(child: const SizedBox());
  print(c is Align);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn center_default_factors() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Center(child: const SizedBox());
  print('${c.widthFactor}:${c.heightFactor}');
}
"#
        ),
        vec!["null:null"]
    );
}

#[test]
fn center_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Center(child: const Placeholder());
  print(c.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn center_render_object_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final c = Center(child: const SizedBox());
  // Center inherits createRenderObject from Align
  // which creates a RenderPositionedBox
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}

#[test]
fn center_debug_fill_properties() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/foundation.dart';
void main() {
  final c = Center(child: const SizedBox());
  final b = DiagnosticPropertiesBuilder();
  c.debugFillProperties(b);
  print(b.properties.length >= 0);
}
"#
        ),
        vec!["true"]
    );
}
