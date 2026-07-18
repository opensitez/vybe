use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material CircularProgressIndicator
// ═══════════════════════════════════════════════════════════

#[test]
fn circular_progress_indicator_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const cpi = CircularProgressIndicator();
  print(cpi is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn circular_progress_indicator_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const cpi = CircularProgressIndicator(value: 0.75);
  print(cpi.value);
}
"#
        ),
        vec!["0.75"]
    );
}

#[test]
fn circular_progress_indicator_background_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const cpi = CircularProgressIndicator(backgroundColor: Color(0xFF0000FF));
  print(cpi.backgroundColor?.value == 0xFF0000FF);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn circular_progress_indicator_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const cpi = CircularProgressIndicator(color: Color(0xFFFF0000));
  print(cpi.color?.value == 0xFFFF0000);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn circular_progress_indicator_stroke_width() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const cpi = CircularProgressIndicator(strokeWidth: 8.0);
  print(cpi.strokeWidth);
}
"#
        ),
        vec!["8.0"]
    );
}

#[test]
fn circular_progress_indicator_stroke_align() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const cpi = CircularProgressIndicator(strokeAlign: 1.0);
  print(cpi.strokeAlign);
}
"#
        ),
        vec!["1.0"]
    );
}
