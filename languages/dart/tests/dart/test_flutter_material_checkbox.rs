use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Checkbox
// ═══════════════════════════════════════════════════════════

#[test]
fn checkbox_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final cb = Checkbox(value: true, onChanged: (bool? value) {});
  print(cb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn checkbox_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final cb = Checkbox(value: false, onChanged: (bool? value) {});
  print(cb.value);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn checkbox_tristate() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final cb = Checkbox(value: null, tristate: true, onChanged: (bool? value) {});
  print('${cb.value}:${cb.tristate}');
}
"#
        ),
        vec!["null:true"]
    );
}

#[test]
fn checkbox_active_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final cb = Checkbox(
    value: true,
    activeColor: const Color(0xFF00FF00),
    onChanged: (bool? value) {},
  );
  print(cb.activeColor?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn checkbox_check_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final cb = Checkbox(
    value: true,
    checkColor: const Color(0xFF112233),
    onChanged: (bool? value) {},
  );
  print(cb.checkColor?.value == 0xFF112233);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn checkbox_is_error() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final cb = Checkbox(
    value: true,
    isError: true,
    onChanged: (bool? value) {},
  );
  print(cb.isError);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn checkbox_focus_node() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fn = FocusNode();
  final cb = Checkbox(
    value: true,
    focusNode: fn,
    onChanged: (bool? value) {},
  );
  print(cb.focusNode == fn);
}
"#
        ),
        vec!["true"]
    );
}
