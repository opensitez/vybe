use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets RichText
// ═══════════════════════════════════════════════════════════

#[test]
fn rich_text_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rt = RichText(text: const TextSpan(text: 'Hello'));
  print(rt.text.toPlainText());
}
"#
        ),
        vec!["Hello"]
    );
}

#[test]
fn rich_text_multiple_spans() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rt = RichText(
    text: const TextSpan(
      children: [
        TextSpan(text: 'A'),
        TextSpan(text: 'B'),
      ]
    ),
  );
  print(rt.text.toPlainText());
}
"#
        ),
        vec!["AB"]
    );
}

#[test]
fn rich_text_align() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rt = RichText(
    text: const TextSpan(text: 'Center'),
    textAlign: TextAlign.center,
  );
  print(rt.textAlign == TextAlign.center);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn rich_text_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rt = RichText(
    text: const TextSpan(text: 'RTL'),
    textDirection: TextDirection.rtl,
  );
  print(rt.textDirection == TextDirection.rtl);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn rich_text_soft_wrap() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rt = RichText(
    text: const TextSpan(text: 'Wrap'),
    softWrap: false,
  );
  print(rt.softWrap);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn rich_text_overflow() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rt = RichText(
    text: const TextSpan(text: 'Overflow'),
    overflow: TextOverflow.fade,
  );
  print(rt.overflow == TextOverflow.fade);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn rich_text_max_lines() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rt = RichText(
    text: const TextSpan(text: 'Lines'),
    maxLines: 3,
  );
  print(rt.maxLines);
}
"#
        ),
        vec!["3"]
    );
}

#[test]
fn rich_text_is_multi_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rt = RichText(text: const TextSpan(text: 'A'));
  print(rt is MultiChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn rich_text_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final rt = RichText(text: const TextSpan(text: 'RenderParagraph'));
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
