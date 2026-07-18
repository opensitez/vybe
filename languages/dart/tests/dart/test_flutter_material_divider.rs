use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Divider
// ═══════════════════════════════════════════════════════════

#[test]
fn divider_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const d = Divider();
  print(d is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn divider_height() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const d = Divider(height: 20.0);
  print(d.height);
}
"#
        ),
        vec!["20.0"]
    );
}

#[test]
fn divider_thickness() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const d = Divider(thickness: 2.0);
  print(d.thickness);
}
"#
        ),
        vec!["2.0"]
    );
}

#[test]
fn divider_indent() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const d = Divider(indent: 10.0);
  print(d.indent);
}
"#
        ),
        vec!["10.0"]
    );
}

#[test]
fn divider_end_indent() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const d = Divider(endIndent: 15.0);
  print(d.endIndent);
}
"#
        ),
        vec!["15.0"]
    );
}

#[test]
fn divider_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const d = Divider(color: Color(0xFF123456));
  print(d.color?.value == 0xFF123456);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn vertical_divider_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const vd = VerticalDivider(width: 10.0, thickness: 2.0);
  print('${vd.width}:${vd.thickness}');
}
"#
        ),
        vec!["10.0:2.0"]
    );
}
