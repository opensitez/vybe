use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material TextField
// ═══════════════════════════════════════════════════════════

#[test]
fn text_field_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tf = TextField();
  print(tf is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_field_controller() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final c = TextEditingController(text: 'Hello');
  final tf = TextField(controller: c);
  print(tf.controller?.text);
}
"#
        ),
        vec!["Hello"]
    );
}

#[test]
fn text_field_focus_node() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fn = FocusNode();
  final tf = TextField(focusNode: fn);
  print(tf.focusNode == fn);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_field_decoration() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tf = TextField(decoration: InputDecoration(labelText: 'Label'));
  print(tf.decoration?.labelText);
}
"#
        ),
        vec!["Label"]
    );
}

#[test]
fn text_field_keyboard_type() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tf = TextField(keyboardType: TextInputType.number);
  print(tf.keyboardType == TextInputType.number);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_field_obscure_text() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tf = TextField(obscureText: true);
  print(tf.obscureText);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_field_max_lines() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tf = TextField(maxLines: 5);
  print(tf.maxLines);
}
"#
        ),
        vec!["5"]
    );
}

#[test]
fn text_field_on_changed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tf = TextField(onChanged: (String value) {});
  print(tf.onChanged != null);
}
"#
        ),
        vec!["true"]
    );
}
