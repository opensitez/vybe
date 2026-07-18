use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Positioned
// ═══════════════════════════════════════════════════════════

#[test]
fn positioned_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Positioned(left: 10.0, top: 20.0, child: const SizedBox());
  print('${p.left}:${p.top}');
}
"#
        ),
        vec!["10.0:20.0"]
    );
}

#[test]
fn positioned_right_bottom() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Positioned(right: 30.0, bottom: 40.0, child: const SizedBox());
  print('${p.right}:${p.bottom}');
}
"#
        ),
        vec!["30.0:40.0"]
    );
}

#[test]
fn positioned_width_height() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Positioned(width: 100.0, height: 200.0, child: const SizedBox());
  print('${p.width}:${p.height}');
}
"#
        ),
        vec!["100.0:200.0"]
    );
}

#[test]
fn positioned_fill() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Positioned.fill(child: const SizedBox());
  print('${p.left}:${p.top}:${p.right}:${p.bottom}');
}
"#
        ),
        vec!["0.0:0.0:0.0:0.0"]
    );
}

#[test]
fn positioned_fill_with_margins() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Positioned.fill(left: 10.0, child: const SizedBox());
  print('${p.left}:${p.right}');
}
"#
        ),
        vec!["10.0:0.0"]
    );
}

#[test]
fn positioned_from_rect() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
void main() {
  final rect = Rect.fromLTRB(10.0, 20.0, 30.0, 40.0);
  final p = Positioned.fromRect(rect: rect, child: const SizedBox());
  print('${p.left}:${p.top}:${p.width}:${p.height}');
}
"#
        ),
        vec!["10.0:20.0:20.0:20.0"]
    );
}

#[test]
fn positioned_from_relative_rect() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rr = RelativeRect.fromLTRB(10.0, 20.0, 30.0, 40.0);
  final p = Positioned.fromRelativeRect(rect: rr, child: const SizedBox());
  print('${p.left}:${p.top}:${p.right}:${p.bottom}');
}
"#
        ),
        vec!["10.0:20.0:30.0:40.0"]
    );
}

#[test]
fn positioned_directional() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = PositionedDirectional(start: 10.0, end: 20.0, child: const SizedBox());
  print('${p.start}:${p.end}');
}
"#
        ),
        vec!["10.0:20.0"]
    );
}

#[test]
fn positioned_directional_is_parent_data_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = PositionedDirectional(child: const SizedBox());
  print(p is ParentDataWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn positioned_is_parent_data_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final p = Positioned(child: const SizedBox());
  print(p is ParentDataWidget);
}
"#
        ),
        vec!["true"]
    );
}
