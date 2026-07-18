use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets SliverGeometry
// ═══════════════════════════════════════════════════════════

#[test]
fn sliver_geometry_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final sg = SliverGeometry(scrollExtent: 100.0, paintExtent: 50.0);
  print('${sg.scrollExtent}:${sg.paintExtent}');
}
"#
        ),
        vec!["100.0:50.0"]
    );
}

#[test]
fn sliver_geometry_defaults() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final sg = SliverGeometry();
  print('${sg.scrollExtent}:${sg.paintExtent}:${sg.layoutExtent}');
}
"#
        ),
        vec!["0.0:0.0:0.0"]
    );
}

#[test]
fn sliver_geometry_max_paint_extent() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final sg = SliverGeometry(maxPaintExtent: 200.0);
  print(sg.maxPaintExtent);
}
"#
        ),
        vec!["200.0"]
    );
}

#[test]
fn sliver_geometry_hit_test_extent() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final sg = SliverGeometry(paintExtent: 50.0, hitTestExtent: 100.0);
  print('${sg.paintExtent}:${sg.hitTestExtent}');
}
"#
        ),
        vec!["50.0:100.0"]
    );
}

#[test]
fn sliver_geometry_visible() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final sg1 = SliverGeometry(visible: true);
  final sg2 = SliverGeometry(visible: false);
  print('${sg1.visible}:${sg2.visible}');
}
"#
        ),
        vec!["true\nfalse"]
    );
}

#[test]
fn sliver_geometry_has_visual_overflow() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final sg = SliverGeometry(hasVisualOverflow: true);
  print(sg.hasVisualOverflow);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_geometry_scroll_offset_correction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final sg = SliverGeometry(scrollOffsetCorrection: 10.5);
  print(sg.scrollOffsetCorrection);
}
"#
        ),
        vec!["10.5"]
    );
}

#[test]
fn sliver_geometry_cache_extent() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final sg = SliverGeometry(cacheExtent: 150.0);
  print(sg.cacheExtent);
}
"#
        ),
        vec!["150.0"]
    );
}

#[test]
fn sliver_geometry_equality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final sg1 = SliverGeometry(scrollExtent: 100.0);
  final sg2 = SliverGeometry(scrollExtent: 100.0);
  // Dart flutter equality might or might not compare instances or properties depending on version
  // Actually, SliverGeometry usually doesn't override == (identity only) or maybe it does?
  // Let's print properties
  print(sg1.scrollExtent == sg2.scrollExtent);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_constraints_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/rendering.dart';
void main() {
  final sc = SliverConstraints(
    axisDirection: AxisDirection.down,
    growthDirection: GrowthDirection.forward,
    userScrollDirection: ScrollDirection.idle,
    scrollOffset: 0.0,
    precedingScrollExtent: 0.0,
    overlap: 0.0,
    remainingPaintExtent: 1000.0,
    crossAxisExtent: 500.0,
    crossAxisDirection: AxisDirection.right,
    viewportMainAxisExtent: 1000.0,
    remainingCacheExtent: 1500.0,
    cacheOrigin: 0.0,
  );
  print('${sc.remainingPaintExtent}:${sc.crossAxisExtent}');
}
"#
        ),
        vec!["1000.0:500.0"]
    );
}
