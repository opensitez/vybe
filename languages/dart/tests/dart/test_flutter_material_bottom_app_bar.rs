use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material BottomAppBar
// ═══════════════════════════════════════════════════════════

#[test]
fn bottom_app_bar_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomAppBar();
  print(b is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bottom_app_bar_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomAppBar(color: const Color(0xFF112233));
  print(b.color?.value == 0xFF112233);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bottom_app_bar_elevation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomAppBar(elevation: 4.0);
  print(b.elevation);
}
"#
        ),
        vec!["4.0"]
    );
}

#[test]
fn bottom_app_bar_shape() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomAppBar(shape: const CircularNotchedRectangle());
  print(b.shape is CircularNotchedRectangle);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bottom_app_bar_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomAppBar(clipBehavior: Clip.antiAlias);
  print(b.clipBehavior == Clip.antiAlias);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bottom_app_bar_notch_margin() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomAppBar(notchMargin: 5.0);
  print(b.notchMargin);
}
"#
        ),
        vec!["5.0"]
    );
}

#[test]
fn bottom_app_bar_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomAppBar(child: const Placeholder());
  print(b.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}
