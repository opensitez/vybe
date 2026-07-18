use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: cupertino CupertinoSlider
// ═══════════════════════════════════════════════════════════

#[test]
fn cupertino_slider_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSlider(
    value: 0.5,
    onChanged: (double newValue) {},
  );
  print(cs is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_slider_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSlider(
    value: 0.7,
    onChanged: (double newValue) {},
  );
  print(cs.value);
}
"#
        ),
        vec!["0.7"]
    );
}

#[test]
fn cupertino_slider_min_max() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSlider(
    value: 5.0,
    min: 0.0,
    max: 10.0,
    onChanged: (double newValue) {},
  );
  print('${cs.min}:${cs.max}');
}
"#
        ),
        vec!["0.0:10.0"]
    );
}

#[test]
fn cupertino_slider_divisions() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSlider(
    value: 0.5,
    divisions: 10,
    onChanged: (double newValue) {},
  );
  print(cs.divisions);
}
"#
        ),
        vec!["10"]
    );
}

#[test]
fn cupertino_slider_active_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSlider(
    value: 0.5,
    activeColor: const Color(0xFF00FF00),
    onChanged: (double newValue) {},
  );
  print(cs.activeColor?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_slider_thumb_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSlider(
    value: 0.5,
    thumbColor: const Color(0xFF123456),
    onChanged: (double newValue) {},
  );
  print(cs.thumbColor.value == 0xFF123456);
}
"#
        ),
        vec!["true"]
    );
}
