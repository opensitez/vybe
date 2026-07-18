use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: cupertino CupertinoButton
// ═══════════════════════════════════════════════════════════

#[test]
fn cupertino_button_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cb = CupertinoButton(
    onPressed: () {},
    child: const Text('Button'),
  );
  print(cb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_button_on_pressed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cb = CupertinoButton(
    onPressed: null,
    child: const Text('Disabled'),
  );
  print(cb.onPressed == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_button_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cb = CupertinoButton(
    onPressed: () {},
    child: const Placeholder(),
  );
  print(cb.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_button_filled() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cb = CupertinoButton.filled(
    onPressed: () {},
    child: const Text('Filled'),
  );
  print(cb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_button_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cb = CupertinoButton(
    color: const Color(0xFF00FF00),
    onPressed: () {},
    child: const Text('Color'),
  );
  print(cb.color?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_button_disabled_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cb = CupertinoButton(
    disabledColor: const Color(0xFF555555),
    onPressed: null,
    child: const Text('Disabled'),
  );
  print(cb.disabledColor.value == 0xFF555555);
}
"#
        ),
        vec!["true"]
    );
}
