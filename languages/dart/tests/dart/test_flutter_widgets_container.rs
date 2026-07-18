use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Container
// ═══════════════════════════════════════════════════════════

#[test]
fn container_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container();
  print(c != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn container_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container(color: const Color(0xFFFF0000));
  print(c.color!.value == 0xFFFF0000);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn container_width_height() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container(width: 100.0, height: 200.0);
  // Internally container uses constraints if only width/height are given
  print('${c.constraints?.minWidth}:${c.constraints?.minHeight}');
}
"#
        ),
        vec!["100.0:200.0"]
    );
}

#[test]
fn container_padding() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container(padding: const EdgeInsets.all(8.0));
  print(c.padding is EdgeInsets);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn container_margin() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container(margin: const EdgeInsets.symmetric(horizontal: 10.0));
  print(c.margin is EdgeInsets);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn container_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container(alignment: Alignment.center);
  print(c.alignment == Alignment.center);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn container_decoration() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container(decoration: const BoxDecoration(color: Color(0xFF00FF00)));
  print(c.decoration != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn container_foreground_decoration() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container(foregroundDecoration: const BoxDecoration(color: Color(0xFF00FF00)));
  print(c.foregroundDecoration != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn container_constraints() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container(constraints: const BoxConstraints(minWidth: 50.0));
  print(c.constraints?.minWidth);
}
"#
        ),
        vec!["50.0"]
    );
}

#[test]
fn container_transform() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container(transform: Matrix4.identity());
  print(c.transform != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn container_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Container(clipBehavior: Clip.antiAlias);
  print(c.clipBehavior == Clip.antiAlias);
}
"#
        ),
        vec!["true"]
    );
}
