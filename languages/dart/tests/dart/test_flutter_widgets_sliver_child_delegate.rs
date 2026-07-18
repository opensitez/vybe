use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets SliverChildDelegate
// ═══════════════════════════════════════════════════════════

#[test]
fn sliver_child_builder_delegate_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final delegate = SliverChildBuilderDelegate((context, index) => const SizedBox(), childCount: 10);
  print(delegate.estimatedChildCount);
}
"#
        ),
        vec!["10"]
    );
}

#[test]
fn sliver_child_builder_delegate_build() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final delegate = SliverChildBuilderDelegate((context, index) => const Placeholder());
  final w = const SizedBox();
  final e = w.createElement();
  final child = delegate.build(e, 0);
  print(child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_child_builder_delegate_should_rebuild() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final delegate1 = SliverChildBuilderDelegate((context, index) => const SizedBox());
  final delegate2 = SliverChildBuilderDelegate((context, index) => const SizedBox());
  print(delegate1.shouldRebuild(delegate2));
}
"#
        ),
        vec!["true"] // usually true because it's a new instance with a new closure
    );
}

#[test]
fn sliver_child_list_delegate_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final delegate = SliverChildListDelegate([const SizedBox(), const SizedBox()]);
  print(delegate.estimatedChildCount);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn sliver_child_list_delegate_build() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w1 = const SizedBox();
  final w2 = const Placeholder();
  final delegate = SliverChildListDelegate([w1, w2]);
  final e = const SizedBox().createElement();
  print(delegate.build(e, 1) == w2);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_child_list_delegate_out_of_bounds() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final delegate = SliverChildListDelegate([const SizedBox()]);
  final e = const SizedBox().createElement();
  print(delegate.build(e, 1) == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_child_list_delegate_should_rebuild() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final delegate1 = SliverChildListDelegate([const SizedBox()]);
  final delegate2 = SliverChildListDelegate([const SizedBox()]);
  print(delegate1.shouldRebuild(delegate2));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_child_builder_delegate_find_index_by_key() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final delegate = SliverChildBuilderDelegate((context, index) => const SizedBox(), findChildIndexCallback: (key) {
    if (key == const ValueKey(1)) return 5;
    return null;
  });
  print(delegate.findIndexByKey(const ValueKey(1)));
}
"#
        ),
        vec!["5"]
    );
}

#[test]
fn sliver_child_list_delegate_find_index_by_key() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final delegate = SliverChildListDelegate([
    SizedBox(key: ValueKey('a')),
    SizedBox(key: ValueKey('b')),
  ]);
  print(delegate.findIndexByKey(ValueKey('b')));
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn sliver_child_delegate_estimate_max_scroll_offset() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final delegate = SliverChildListDelegate([const SizedBox()]);
  // Usually returns null if not specified or calculatable
  print(delegate.estimateMaxScrollOffset(0, 1, 100, 100) == null);
}
"#
        ),
        vec!["true"]
    );
}
