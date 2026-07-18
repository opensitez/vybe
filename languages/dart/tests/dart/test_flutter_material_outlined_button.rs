use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material OutlinedButton
// ═══════════════════════════════════════════════════════════

#[test]
fn outlined_button_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ob = OutlinedButton(
    onPressed: () {},
    child: const Text('Outlined'),
  );
  print(ob is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn outlined_button_on_pressed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ob = OutlinedButton(
    onPressed: null,
    child: const Text('Disabled'),
  );
  print(ob.enabled);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn outlined_button_icon_constructor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ob = OutlinedButton.icon(
    onPressed: () {},
    icon: const Icon(Icons.download),
    label: const Text('Download'),
  );
  print(ob is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn outlined_button_style() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ob = OutlinedButton(
    style: OutlinedButton.styleFrom(
      side: const BorderSide(color: Color(0xFF123456), width: 2.0),
    ),
    onPressed: () {},
    child: const Text('Style'),
  );
  print(ob.style != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn outlined_button_focus_node() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fn = FocusNode();
  final ob = OutlinedButton(
    focusNode: fn,
    onPressed: () {},
    child: const Text('Focus'),
  );
  print(ob.focusNode == fn);
}
"#
        ),
        vec!["true"]
    );
}
