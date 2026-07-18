use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Dismissible
// ═══════════════════════════════════════════════════════════

#[test]
fn dismissible_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Dismissible(
    key: const ValueKey('key'),
    child: const SizedBox(),
  );
  print(d is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dismissible_key() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Dismissible(
    key: const ValueKey('my_key'),
    child: const SizedBox(),
  );
  print((d.key as ValueKey).value);
}
"#
        ),
        vec!["my_key"]
    );
}

#[test]
fn dismissible_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Dismissible(
    key: const ValueKey('key'),
    child: const Text('Child'),
  );
  print(d.child is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dismissible_background() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Dismissible(
    key: const ValueKey('key'),
    background: const Placeholder(),
    child: const SizedBox(),
  );
  print(d.background is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dismissible_secondary_background() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Dismissible(
    key: const ValueKey('key'),
    secondaryBackground: const Placeholder(),
    child: const SizedBox(),
  );
  print(d.secondaryBackground is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dismissible_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Dismissible(
    key: const ValueKey('key'),
    direction: DismissDirection.up,
    child: const SizedBox(),
  );
  print(d.direction == DismissDirection.up);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dismissible_on_dismissed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Dismissible(
    key: const ValueKey('key'),
    onDismissed: (DismissDirection direction) {},
    child: const SizedBox(),
  );
  print(d.onDismissed != null);
}
"#
        ),
        vec!["true"]
    );
}
