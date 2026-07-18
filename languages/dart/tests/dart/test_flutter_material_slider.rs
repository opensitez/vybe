use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Slider
// ═══════════════════════════════════════════════════════════

#[test]
fn slider_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Slider(
    value: 0.5,
    onChanged: (double newValue) {},
  );
  print(s is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn slider_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Slider(
    value: 0.3,
    onChanged: (double newValue) {},
  );
  print(s.value);
}
"#
        ),
        vec!["0.3"]
    );
}

#[test]
fn slider_min_max() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Slider(
    value: 50.0,
    min: 10.0,
    max: 100.0,
    onChanged: (double newValue) {},
  );
  print('${s.min}:${s.max}');
}
"#
        ),
        vec!["10.0:100.0"]
    );
}

#[test]
fn slider_divisions() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Slider(
    value: 0.5,
    divisions: 5,
    onChanged: (double newValue) {},
  );
  print(s.divisions);
}
"#
        ),
        vec!["5"]
    );
}

#[test]
fn slider_active_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Slider(
    value: 0.5,
    activeColor: const Color(0xFF00FF00),
    onChanged: (double newValue) {},
  );
  print(s.activeColor?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn slider_inactive_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Slider(
    value: 0.5,
    inactiveColor: const Color(0xFF112233),
    onChanged: (double newValue) {},
  );
  print(s.inactiveColor?.value == 0xFF112233);
}
"#
        ),
        vec!["true"]
    );
}
