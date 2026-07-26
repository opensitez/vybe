#![allow(non_snake_case)]
use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets BoxConstraints
// ═══════════════════════════════════════════════════════════

#[test]
fn box_constraints_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final c = BoxConstraints(minWidth: 10, maxWidth: 100, minHeight: 20, maxHeight: 200);
  print('${c.minWidth}:${c.maxWidth}:${c.minHeight}:${c.maxHeight}');
}
"#
        ),
        vec!["10.0:100.0:20.0:200.0"]
    );
}

#[test]
fn box_constraints_tight() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
import 'dart:ui';
void main() {
  final c = BoxConstraints.tight(Size(50, 50));
  print('${c.minWidth}:${c.maxWidth}:${c.minHeight}:${c.maxHeight}');
  print(c.isTight);
}
"#
        ),
        vec!["50.0:50.0:50.0:50.0\ntrue"]
    );
}

#[test]
fn box_constraints_loose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
import 'dart:ui';
void main() {
  final c = BoxConstraints.loose(Size(100, 100));
  print('${c.minWidth}:${c.maxWidth}:${c.minHeight}:${c.maxHeight}');
}
"#
        ),
        vec!["0.0:100.0:0.0:100.0"]
    );
}

#[test]
fn box_constraints_expand() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final c = BoxConstraints.expand(width: 200, height: 300);
  print('${c.minWidth}:${c.maxWidth}:${c.minHeight}:${c.maxHeight}');
}
"#
        ),
        vec!["200.0:200.0:300.0:300.0"]
    );
}

#[test]
fn box_constraints_tightFor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final c = BoxConstraints.tightFor(width: 50);
  // height is 0 to infinity
  print('${c.minWidth}:${c.maxWidth}:${c.minHeight}:${c.maxHeight == double.infinity}');
}
"#
        ),
        vec!["50.0:50.0:0.0:true"]
    );
}

#[test]
fn box_constraints_constrain() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
import 'dart:ui';
void main() {
  final c = BoxConstraints(minWidth: 50, maxWidth: 100, minHeight: 50, maxHeight: 100);
  final s1 = c.constrain(Size(10, 10)); // too small
  final s2 = c.constrain(Size(200, 200)); // too large
  print('${s1.width}:${s1.height} ${s2.width}:${s2.height}');
}
"#
        ),
        vec!["50.0:50.0 100.0:100.0"]
    );
}

#[test]
fn box_constraints_constrain_width() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final c = BoxConstraints(minWidth: 50, maxWidth: 100);
  print(c.constrainWidth(10));
  print(c.constrainWidth(200));
}
"#
        ),
        vec!["50.0\n100.0"]
    );
}

#[test]
fn box_constraints_constrain_height() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final c = BoxConstraints(minHeight: 50, maxHeight: 100);
  print(c.constrainHeight(10));
  print(c.constrainHeight(200));
}
"#
        ),
        vec!["50.0\n100.0"]
    );
}

#[test]
fn box_constraints_isNormalized() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final c1 = BoxConstraints(minWidth: 100, maxWidth: 50); // invalid
  final c2 = BoxConstraints(minWidth: 50, maxWidth: 100); // valid
  print(c1.isNormalized);
  print(c2.isNormalized);
}
"#
        ),
        vec!["false\ntrue"]
    );
}

#[test]
fn box_constraints_normalize() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final c1 = BoxConstraints(minWidth: 100, maxWidth: 50);
  final norm = c1.normalize();
  print('${norm.minWidth}:${norm.maxWidth}');
}
"#
        ),
        vec!["100.0:100.0"]
    );
}

#[test]
fn box_constraints_deflate() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
import 'package:flutter/painting.dart';
void main() {
  final c = BoxConstraints(minWidth: 100, maxWidth: 200, minHeight: 100, maxHeight: 200);
  final insets = EdgeInsets.all(10);
  final deflated = c.deflate(insets);
  print('${deflated.minWidth}:${deflated.maxWidth}');
}
"#
        ),
        vec!["80.0:180.0"]
    );
}
