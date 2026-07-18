use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material ElevatedButton
// ═══════════════════════════════════════════════════════════

#[test]
fn elevated_button_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final eb = ElevatedButton(
    onPressed: () {},
    child: const Text('Button'),
  );
  print(eb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn elevated_button_on_pressed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final eb = ElevatedButton(
    onPressed: null,
    child: const Text('Disabled'),
  );
  print(eb.onPressed == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn elevated_button_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final eb = ElevatedButton(
    onPressed: () {},
    child: const Placeholder(),
  );
  print(eb.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn elevated_button_icon_constructor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final eb = ElevatedButton.icon(
    onPressed: () {},
    icon: const Icon(Icons.add),
    label: const Text('Add'),
  );
  print(eb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn elevated_button_style() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final eb = ElevatedButton(
    style: ElevatedButton.styleFrom(
      backgroundColor: const Color(0xFF000000),
      elevation: 5.0,
    ),
    onPressed: () {},
    child: const Text('Style'),
  );
  print(eb.style != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn elevated_button_focus_node() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fn = FocusNode();
  final eb = ElevatedButton(
    focusNode: fn,
    onPressed: () {},
    child: const Text('Focus'),
  );
  print(eb.focusNode == fn);
}
"#
        ),
        vec!["true"]
    );
}
