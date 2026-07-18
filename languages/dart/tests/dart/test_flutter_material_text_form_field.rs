use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material TextFormField
// ═══════════════════════════════════════════════════════════

#[test]
fn text_form_field_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tff = TextFormField();
  print(tff is FormField<String>);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_form_field_initial_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tff = TextFormField(initialValue: 'Test');
  print(tff.initialValue);
}
"#
        ),
        vec!["Test"]
    );
}

#[test]
fn text_form_field_controller() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final c = TextEditingController(text: 'Controller');
  final tff = TextFormField(controller: c);
  print(tff.controller?.text);
}
"#
        ),
        vec!["Controller"]
    );
}

#[test]
fn text_form_field_validator() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tff = TextFormField(
    validator: (String? value) {
      if (value == null || value.isEmpty) return 'Error';
      return null;
    },
  );
  print(tff.validator != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_form_field_on_saved() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tff = TextFormField(onSaved: (String? value) {});
  print(tff.onSaved != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_form_field_decoration() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tff = TextFormField(decoration: InputDecoration(hintText: 'Hint'));
  print(tff.decoration?.hintText);
}
"#
        ),
        vec!["Hint"]
    );
}

#[test]
fn text_form_field_obscure_text() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tff = TextFormField(obscureText: true);
  print(tff.obscureText);
}
"#
        ),
        vec!["true"]
    );
}
