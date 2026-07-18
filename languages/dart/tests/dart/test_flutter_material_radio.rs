use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Radio
// ═══════════════════════════════════════════════════════════

#[test]
fn radio_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final r = Radio<int>(
    value: 1,
    groupValue: 1,
    onChanged: (int? newValue) {},
  );
  print(r is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn radio_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final r = Radio<String>(
    value: 'A',
    groupValue: 'B',
    onChanged: (String? newValue) {},
  );
  print('${r.value}:${r.groupValue}');
}
"#
        ),
        vec!["A:B"]
    );
}

#[test]
fn radio_active_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final r = Radio<int>(
    value: 1,
    groupValue: 1,
    activeColor: const Color(0xFF00FF00),
    onChanged: (int? newValue) {},
  );
  print(r.activeColor?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn radio_toggleable() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final r = Radio<int>(
    value: 1,
    groupValue: 1,
    toggleable: true,
    onChanged: (int? newValue) {},
  );
  print(r.toggleable);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn radio_focus_node() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fn = FocusNode();
  final r = Radio<int>(
    value: 1,
    groupValue: 1,
    focusNode: fn,
    onChanged: (int? newValue) {},
  );
  print(r.focusNode == fn);
}
"#
        ),
        vec!["true"]
    );
}
