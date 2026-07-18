use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material BottomSheet
// ═══════════════════════════════════════════════════════════

#[test]
fn bottom_sheet_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final bs = BottomSheet(
    onClosing: () {},
    builder: (context) => const SizedBox(),
  );
  print(bs is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bottom_sheet_elevation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final bs = BottomSheet(
    elevation: 8.0,
    onClosing: () {},
    builder: (context) => const SizedBox(),
  );
  print(bs.elevation);
}
"#
        ),
        vec!["8.0"]
    );
}

#[test]
fn bottom_sheet_enable_drag() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final bs = BottomSheet(
    enableDrag: false,
    onClosing: () {},
    builder: (context) => const SizedBox(),
  );
  print(bs.enableDrag);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn bottom_sheet_on_drag_start() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final bs = BottomSheet(
    onDragStart: (details) {},
    onClosing: () {},
    builder: (context) => const SizedBox(),
  );
  print(bs.onDragStart != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bottom_sheet_animation_controller() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final bs = BottomSheet(
    animationController: AnimationController(
      vsync: const TestVSync(),
      duration: const Duration(seconds: 1),
    ),
    onClosing: () {},
    builder: (context) => const SizedBox(),
  );
  print(bs.animationController != null);
}

class TestVSync implements TickerProvider {
  const TestVSync();
  @override
  Ticker createTicker(TickerCallback onTick) => Ticker(onTick);
}
"#
        ),
        vec!["true"]
    );
}
