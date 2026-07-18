use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material InputDecorator
// ═══════════════════════════════════════════════════════════

#[test]
fn input_decorator_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final id = InputDecorator(
    decoration: const InputDecoration(labelText: 'Label'),
    child: const SizedBox(),
  );
  print(id is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn input_decorator_decoration() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final id = InputDecorator(
    decoration: const InputDecoration(hintText: 'Hint'),
    child: const SizedBox(),
  );
  print(id.decoration.hintText);
}
"#
        ),
        vec!["Hint"]
    );
}

#[test]
fn input_decorator_base_style() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final id = InputDecorator(
    decoration: const InputDecoration(),
    baseStyle: const TextStyle(color: Color(0xFF00FF00)),
    child: const SizedBox(),
  );
  print(id.baseStyle?.color?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn input_decorator_is_focused() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final id = InputDecorator(
    decoration: const InputDecoration(),
    isFocused: true,
    child: const SizedBox(),
  );
  print(id.isFocused);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn input_decorator_is_hovering() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final id = InputDecorator(
    decoration: const InputDecoration(),
    isHovering: true,
    child: const SizedBox(),
  );
  print(id.isHovering);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn input_decorator_expands() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final id = InputDecorator(
    decoration: const InputDecoration(),
    expands: true,
    child: const SizedBox(),
  );
  print(id.expands);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn input_decorator_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final id = InputDecorator(
    decoration: const InputDecoration(),
    child: const Text('Child'),
  );
  print(id.child is Text);
}
"#
        ),
        vec!["true"]
    );
}
