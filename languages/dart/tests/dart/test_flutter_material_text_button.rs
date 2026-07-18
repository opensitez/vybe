use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material TextButton
// ═══════════════════════════════════════════════════════════

#[test]
fn text_button_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tb = TextButton(
    onPressed: () {},
    child: const Text('TextButton'),
  );
  print(tb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_button_on_pressed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tb = TextButton(
    onPressed: null,
    child: const Text('Disabled'),
  );
  print(tb.enabled);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn text_button_icon_constructor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tb = TextButton.icon(
    onPressed: () {},
    icon: const Icon(Icons.share),
    label: const Text('Share'),
  );
  print(tb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_button_style() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tb = TextButton(
    style: TextButton.styleFrom(
      foregroundColor: const Color(0xFF00FF00),
    ),
    onPressed: () {},
    child: const Text('Style'),
  );
  print(tb.style != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_button_autofocus() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tb = TextButton(
    autofocus: true,
    onPressed: () {},
    child: const Text('Autofocus'),
  );
  print(tb.autofocus);
}
"#
        ),
        vec!["true"]
    );
}
