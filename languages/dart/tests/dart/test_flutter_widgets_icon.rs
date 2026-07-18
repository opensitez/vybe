use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Icon
// ═══════════════════════════════════════════════════════════

#[test]
fn icon_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final i = Icon(const IconData(0xe000, fontFamily: 'MaterialIcons'));
  print(i.icon?.codePoint);
}
"#
        ),
        vec!["57344"]
    );
}

#[test]
fn icon_size() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final i = Icon(null, size: 24.0);
  print(i.size);
}
"#
        ),
        vec!["24.0"]
    );
}

#[test]
fn icon_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final i = Icon(null, color: const Color(0xFF00FF00));
  print(i.color?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn icon_semantic_label() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final i = Icon(null, semanticLabel: 'Back');
  print(i.semanticLabel);
}
"#
        ),
        vec!["Back"]
    );
}

#[test]
fn icon_text_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final i = Icon(null, textDirection: TextDirection.rtl);
  print(i.textDirection == TextDirection.rtl);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn icon_data_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final id = const IconData(0xe123, fontFamily: 'CustomIcons');
  print('${id.codePoint}:${id.fontFamily}');
}
"#
        ),
        vec!["57635:CustomIcons"]
    );
}

#[test]
fn icon_data_equality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final id1 = const IconData(0xe123, fontFamily: 'Font');
  final id2 = const IconData(0xe123, fontFamily: 'Font');
  print(id1 == id2);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn icon_is_stateless_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final i = Icon(null);
  print(i is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn icon_theme_data() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final itd = const IconThemeData(color: Color(0xFF112233), size: 30.0);
  print('${itd.color?.value == 0xFF112233}:${itd.size}');
}
"#
        ),
        vec!["true:30.0"]
    );
}

#[test]
fn icon_theme_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final it = IconTheme(
    data: const IconThemeData(size: 40.0),
    child: const SizedBox(),
  );
  print(it.data.size);
}
"#
        ),
        vec!["40.0"]
    );
}
