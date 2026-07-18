use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: animation Curve Functions
// ═══════════════════════════════════════════════════════════

#[test]
fn curves_linear() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  print(Curves.linear.transform(0.5));
}
"#
        ),
        vec!["0.5"]
    );
}

#[test]
fn curves_ease() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final val = Curves.ease.transform(0.5);
  // Standard ease function
  print(val > 0.0 && val < 1.0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn curves_easeIn() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final val = Curves.easeIn.transform(0.5);
  // Ease in accelerates, so at 0.5 it's less than 0.5
  print(val < 0.5);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn curves_easeOut() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final val = Curves.easeOut.transform(0.5);
  // Ease out decelerates, so at 0.5 it's greater than 0.5
  print(val > 0.5);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn curves_easeInOut() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final val = Curves.easeInOut.transform(0.5);
  // Usually around 0.5
  print(val == 0.5);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn curves_bounceIn() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final val = Curves.bounceIn.transform(0.2);
  print(val > 0.0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn curves_elasticIn() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final val = Curves.elasticIn.transform(0.5);
  // Elastic can dip below 0
  print(val != 0.5);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn curves_elasticOut() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final val = Curves.elasticOut.transform(0.5);
  // Elastic out can overshoot 1
  print(val != 0.5);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_curve() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
class MyCurve extends Curve {
  @override
  double transformInternal(double t) {
    return t * t;
  }
}
void main() {
  final c = MyCurve();
  print(c.transform(0.5));
}
"#
        ),
        vec!["0.25"]
    );
}

#[test]
fn flipped_curve() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final flipped = FlippedCurve(Curves.easeIn);
  // Same as easeOut shape but exact values depend on math
  print(flipped.transform(0.5) > 0.5);
}
"#
        ),
        vec!["true"] // 1.0 - easeIn(1.0 - 0.5) = 1.0 - easeIn(0.5). easeIn(0.5) < 0.5. 1.0 - <0.5 is > 0.5
    );
}

#[test]
fn threshold_curve() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final curve = Threshold(0.6);
  print('${curve.transform(0.5)}:${curve.transform(0.7)}');
}
"#
        ),
        vec!["0.0:1.0"]
    );
}

#[test]
fn interval_curve() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/animation.dart';
void main() {
  final curve = Interval(0.2, 0.8, curve: Curves.linear);
  print('${curve.transform(0.1)}:${curve.transform(0.5)}:${curve.transform(0.9)}');
}
"#
        ),
        vec!["0.0:0.5:1.0"] // 0.5 is exactly halfway between 0.2 and 0.8
    );
}
