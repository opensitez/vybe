use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material InkWell
// ═══════════════════════════════════════════════════════════

#[test]
fn ink_well_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final iw = InkWell(child: const SizedBox());
  print(iw is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn ink_well_on_tap() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final iw = InkWell(onTap: () {}, child: const SizedBox());
  print(iw.onTap != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn ink_well_on_double_tap() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final iw = InkWell(onDoubleTap: () {}, child: const SizedBox());
  print(iw.onDoubleTap != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn ink_well_on_long_press() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final iw = InkWell(onLongPress: () {}, child: const SizedBox());
  print(iw.onLongPress != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn ink_well_splash_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final iw = InkWell(
    splashColor: const Color(0xFF00FF00),
    child: const SizedBox(),
  );
  print(iw.splashColor?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn ink_well_highlight_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final iw = InkWell(
    highlightColor: const Color(0xFF123456),
    child: const SizedBox(),
  );
  print(iw.highlightColor?.value == 0xFF123456);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn ink_well_border_radius() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final iw = InkWell(
    borderRadius: BorderRadius.circular(10.0),
    child: const SizedBox(),
  );
  print(iw.borderRadius is BorderRadius);
}
"#
        ),
        vec!["true"]
    );
}
