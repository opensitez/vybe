use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Card
// ═══════════════════════════════════════════════════════════

#[test]
fn card_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final c = Card(child: const SizedBox());
  print(c is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn card_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final c = Card(color: const Color(0xFF000000));
  print(c.color?.value == 0xFF000000);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn card_elevation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final c = Card(elevation: 5.0);
  print(c.elevation);
}
"#
        ),
        vec!["5.0"]
    );
}

#[test]
fn card_shape() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final c = Card(
    shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
  );
  print(c.shape is RoundedRectangleBorder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn card_margin() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final c = Card(margin: const EdgeInsets.all(8.0));
  print((c.margin as EdgeInsets).top);
}
"#
        ),
        vec!["8.0"]
    );
}

#[test]
fn card_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final c = Card(clipBehavior: Clip.antiAlias);
  print(c.clipBehavior == Clip.antiAlias);
}
"#
        ),
        vec!["true"]
    );
}
