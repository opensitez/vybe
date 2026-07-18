use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material LinearProgressIndicator
// ═══════════════════════════════════════════════════════════

#[test]
fn linear_progress_indicator_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const lpi = LinearProgressIndicator();
  print(lpi is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn linear_progress_indicator_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const lpi = LinearProgressIndicator(value: 0.5);
  print(lpi.value);
}
"#
        ),
        vec!["0.5"]
    );
}

#[test]
fn linear_progress_indicator_background_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const lpi = LinearProgressIndicator(backgroundColor: Color(0xFF112233));
  print(lpi.backgroundColor?.value == 0xFF112233);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn linear_progress_indicator_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const lpi = LinearProgressIndicator(color: Color(0xFF00FF00));
  print(lpi.color?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn linear_progress_indicator_min_height() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const lpi = LinearProgressIndicator(minHeight: 10.0);
  print(lpi.minHeight);
}
"#
        ),
        vec!["10.0"]
    );
}

#[test]
fn linear_progress_indicator_value_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final lpi = LinearProgressIndicator(
    valueColor: AlwaysStoppedAnimation<Color>(Color(0xFF556677)),
  );
  print(lpi.valueColor != null);
}
"#
        ),
        vec!["true"]
    );
}
