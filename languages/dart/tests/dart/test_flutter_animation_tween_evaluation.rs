use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: animation Tween Evaluation
// ═══════════════════════════════════════════════════════════

#[test]
fn tween_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final tween = Tween<double>(begin: 0.0, end: 10.0);
  print('${tween.begin}:${tween.end}');
}
"#
        ),
        vec!["0.0:10.0"]
    );
}

#[test]
fn tween_lerp_double() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final tween = Tween<double>(begin: 0.0, end: 10.0);
  print(tween.lerp(0.5));
}
"#
        ),
        vec!["5.0"]
    );
}

#[test]
fn tween_evaluate() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final tween = Tween<double>(begin: 100.0, end: 200.0);
  final anim = AlwaysStoppedAnimation(0.25);
  print(tween.evaluate(anim));
}
"#
        ),
        vec!["125.0"]
    );
}

#[test]
fn tween_transform() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final tween = Tween<double>(begin: 50.0, end: 100.0);
  print(tween.transform(0.5));
}
"#
        ),
        vec!["75.0"]
    );
}

#[test]
fn color_tween() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
import 'dart:ui';
void main() {
  final tween = ColorTween(begin: Color(0xFF000000), end: Color(0xFFFFFFFF));
  final c = tween.lerp(0.5);
  // Color lerping is per-channel, so 0x7F or 0x80 usually
  print(c?.alpha == 255);
  print(c?.red == 127 || c?.red == 128);
}
"#
        ),
        // Normally alpha is 255. Red is halfway between 0 and 255 = 127 or 128
        vec!["true\ntrue"]
    );
}

#[test]
fn rect_tween() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
import 'dart:ui';
void main() {
  final tween = RectTween(
    begin: Rect.fromLTRB(0, 0, 10, 10),
    end: Rect.fromLTRB(10, 10, 30, 30),
  );
  final r = tween.lerp(0.5);
  print('${r?.left}:${r?.right}');
}
"#
        ),
        vec!["5.0:20.0"]
    );
}

#[test]
fn int_tween() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final tween = IntTween(begin: 0, end: 10);
  print(tween.lerp(0.55)); // 5.5 rounds to 6 in IntTween (it uses round())
}
"#
        ),
        vec!["6"]
    );
}

#[test]
fn step_tween() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final tween = StepTween(begin: 0, end: 10);
  // StepTween uses floor() instead of round()
  print(tween.lerp(0.55)); // 5.5 floor is 5
}
"#
        ),
        vec!["5"]
    );
}

#[test]
fn reverse_tween() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final tween = ReverseTween(Tween<double>(begin: 0.0, end: 10.0));
  // reverse tween uses (1.0 - t) for the parent
  print(tween.lerp(0.25)); // parent sees 0.75, so 7.5
}
"#
        ),
        vec!["7.5"]
    );
}

#[test]
fn tween_chain() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final tween = Tween<double>(begin: 0.0, end: 10.0);
  final curve = CurveTween(curve: Curves.easeIn);
  final chained = tween.chain(curve);
  final val = chained.transform(0.5);
  // easeIn at 0.5 is < 0.5, so val < 5.0
  print(val < 5.0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn curve_tween() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final tween = CurveTween(curve: Curves.bounceIn);
  final val = tween.transform(0.5);
  print(val > 0.0 && val < 1.0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn size_tween() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
import 'dart:ui';
void main() {
  final tween = SizeTween(begin: Size(10, 10), end: Size(30, 30));
  final s = tween.lerp(0.5);
  print('${s?.width}:${s?.height}');
}
"#
        ),
        vec!["20.0:20.0"]
    );
}
