use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Stack
// ═══════════════════════════════════════════════════════════

#[test]
fn stack_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Stack(children: [const SizedBox()]);
  print(s.children.length);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn stack_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Stack(alignment: Alignment.bottomRight);
  print(s.alignment == Alignment.bottomRight);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn stack_text_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Stack(textDirection: TextDirection.rtl);
  print(s.textDirection == TextDirection.rtl);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn stack_fit() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Stack(fit: StackFit.expand);
  print(s.fit == StackFit.expand);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn stack_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Stack(clipBehavior: Clip.none);
  print(s.clipBehavior == Clip.none);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn stack_is_multi_child_render_object_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = Stack();
  print(s is MultiChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn indexed_stack_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = IndexedStack(index: 1, children: [const SizedBox(), const SizedBox()]);
  print(s.index);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn indexed_stack_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = IndexedStack(alignment: Alignment.center);
  print(s.alignment == Alignment.center);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn indexed_stack_sizing() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = IndexedStack(sizing: StackFit.loose);
  print(s.sizing == StackFit.loose);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn indexed_stack_is_stack() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final s = IndexedStack();
  print(s is Stack);
}
"#
        ),
        vec!["true"]
    );
}
