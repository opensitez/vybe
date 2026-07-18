use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Text
// ═══════════════════════════════════════════════════════════

#[test]
fn text_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Text('Hello World');
  print(t.data);
}
"#
        ),
        vec!["Hello World"]
    );
}

#[test]
fn text_style() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Text('Styled', style: TextStyle(fontSize: 20.0));
  print(t.style?.fontSize);
}
"#
        ),
        vec!["20.0"]
    );
}

#[test]
fn text_align() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Text('Align', textAlign: TextAlign.center);
  print(t.textAlign == TextAlign.center);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Text('RTL', textDirection: TextDirection.rtl);
  print(t.textDirection == TextDirection.rtl);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_soft_wrap() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Text('Wrap', softWrap: false);
  print(t.softWrap);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn text_overflow() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Text('Overflow', overflow: TextOverflow.ellipsis);
  print(t.overflow == TextOverflow.ellipsis);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_max_lines() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Text('Lines', maxLines: 2);
  print(t.maxLines);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn text_rich() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Text.rich(TextSpan(text: 'Rich'));
  print(t.textSpan != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_is_stateless() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Text('Widget');
  print(t is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn default_text_style() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final dts = DefaultTextStyle(
    style: TextStyle(fontSize: 16.0),
    child: const SizedBox(),
  );
  print(dts.style.fontSize);
}
"#
        ),
        vec!["16.0"]
    );
}
