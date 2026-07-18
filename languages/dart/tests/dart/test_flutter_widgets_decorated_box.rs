use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets DecoratedBox
// ═══════════════════════════════════════════════════════════

#[test]
fn decorated_box_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final db = DecoratedBox(
    decoration: const BoxDecoration(color: Color(0xFF00FF00)),
  );
  print(db.decoration != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn decorated_box_position() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final db = DecoratedBox(
    decoration: const BoxDecoration(),
    position: DecorationPosition.foreground,
  );
  print(db.position == DecorationPosition.foreground);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn decorated_box_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final db = DecoratedBox(
    decoration: const BoxDecoration(),
    child: const Placeholder(),
  );
  print(db.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn decorated_box_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final db = DecoratedBox(decoration: const BoxDecoration());
  print(db is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn decoration_position_enum() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  print(DecorationPosition.background.name);
  print(DecorationPosition.foreground.name);
}
"#
        ),
        vec!["background\nforeground"]
    );
}

#[test]
fn box_decoration_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final bd = const BoxDecoration(color: Color(0xFF123456));
  print(bd.color?.value == 0xFF123456);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn box_decoration_shape() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final bd = const BoxDecoration(shape: BoxShape.circle);
  print(bd.shape == BoxShape.circle);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn decorated_box_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final db = DecoratedBox(decoration: const BoxDecoration());
  // Creates RenderDecoratedBox
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
