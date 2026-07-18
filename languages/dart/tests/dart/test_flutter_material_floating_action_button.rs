use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material FloatingActionButton
// ═══════════════════════════════════════════════════════════

#[test]
fn floating_action_button_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fab = FloatingActionButton(
    onPressed: () {},
    child: const Icon(Icons.add),
  );
  print(fab is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn floating_action_button_on_pressed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fab = FloatingActionButton(
    onPressed: null,
    child: const Icon(Icons.add),
  );
  print(fab.onPressed == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn floating_action_button_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fab = FloatingActionButton(
    onPressed: () {},
    child: const Text('FAB'),
  );
  print(fab.child is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn floating_action_button_tooltip() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fab = FloatingActionButton(
    onPressed: () {},
    tooltip: 'Action',
    child: const Icon(Icons.add),
  );
  print(fab.tooltip);
}
"#
        ),
        vec!["Action"]
    );
}

#[test]
fn floating_action_button_extended() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fab = FloatingActionButton.extended(
    onPressed: () {},
    label: const Text('Extended'),
    icon: const Icon(Icons.add),
  );
  print(fab is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn floating_action_button_elevation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fab = FloatingActionButton(
    onPressed: () {},
    elevation: 10.0,
    child: const Icon(Icons.add),
  );
  print(fab.elevation);
}
"#
        ),
        vec!["10.0"]
    );
}
