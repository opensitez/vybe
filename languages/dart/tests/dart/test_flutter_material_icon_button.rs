use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material IconButton
// ═══════════════════════════════════════════════════════════

#[test]
fn icon_button_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ib = IconButton(
    icon: const Icon(Icons.add),
    onPressed: () {},
  );
  print(ib is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn icon_button_on_pressed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ib = IconButton(
    icon: const Icon(Icons.add),
    onPressed: null,
  );
  print(ib.onPressed == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn icon_button_icon() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ib = IconButton(
    icon: const Text('Icon'),
    onPressed: () {},
  );
  print(ib.icon is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn icon_button_icon_size() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ib = IconButton(
    icon: const Icon(Icons.add),
    iconSize: 32.0,
    onPressed: () {},
  );
  print(ib.iconSize);
}
"#
        ),
        vec!["32.0"]
    );
}

#[test]
fn icon_button_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ib = IconButton(
    icon: const Icon(Icons.add),
    color: const Color(0xFF123456),
    onPressed: () {},
  );
  print(ib.color?.value == 0xFF123456);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn icon_button_tooltip() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ib = IconButton(
    icon: const Icon(Icons.add),
    tooltip: 'Add item',
    onPressed: () {},
  );
  print(ib.tooltip);
}
"#
        ),
        vec!["Add item"]
    );
}
