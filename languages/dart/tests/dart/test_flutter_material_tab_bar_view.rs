use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material TabBarView
// ═══════════════════════════════════════════════════════════

#[test]
fn tab_bar_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tv = TabBarView(children: const [SizedBox()]);
  print(tv is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn tab_bar_view_children() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tv = TabBarView(
    children: const [
      Placeholder(),
      SizedBox(),
    ],
  );
  print(tv.children.length);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn tab_bar_view_physics() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tv = TabBarView(
    physics: const BouncingScrollPhysics(),
    children: const [SizedBox()],
  );
  print(tv.physics is BouncingScrollPhysics);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn tab_bar_view_drag_start_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
import 'package:flutter/gestures.dart';
void main() {
  final tv = TabBarView(
    dragStartBehavior: DragStartBehavior.down,
    children: const [SizedBox()],
  );
  print(tv.dragStartBehavior == DragStartBehavior.down);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn tab_bar_view_viewport_fraction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tv = TabBarView(
    viewportFraction: 0.8,
    children: const [SizedBox()],
  );
  print(tv.viewportFraction);
}
"#
        ),
        vec!["0.8"]
    );
}
