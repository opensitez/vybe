use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Switch
// ═══════════════════════════════════════════════════════════

#[test]
fn switch_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Switch(
    value: true,
    onChanged: (bool newValue) {},
  );
  print(s is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn switch_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Switch(
    value: false,
    onChanged: (bool newValue) {},
  );
  print(s.value);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn switch_active_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Switch(
    value: true,
    activeColor: const Color(0xFF00FF00),
    onChanged: (bool newValue) {},
  );
  print(s.activeColor?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn switch_active_track_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Switch(
    value: true,
    activeTrackColor: const Color(0xFF112233),
    onChanged: (bool newValue) {},
  );
  print(s.activeTrackColor?.value == 0xFF112233);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn switch_inactive_thumb_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Switch(
    value: false,
    inactiveThumbColor: const Color(0xFF445566),
    onChanged: (bool newValue) {},
  );
  print(s.inactiveThumbColor?.value == 0xFF445566);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn switch_inactive_track_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Switch(
    value: false,
    inactiveTrackColor: const Color(0xFF778899),
    onChanged: (bool newValue) {},
  );
  print(s.inactiveTrackColor?.value == 0xFF778899);
}
"#
        ),
        vec!["true"]
    );
}
