use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material SliverAppBar
// ═══════════════════════════════════════════════════════════

#[test]
fn sliver_app_bar_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = SliverAppBar();
  print(s is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_app_bar_title() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = SliverAppBar(title: const Text('Title'));
  print(s.title is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_app_bar_floating() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = SliverAppBar(floating: true);
  print(s.floating);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_app_bar_pinned() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = SliverAppBar(pinned: true);
  print(s.pinned);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_app_bar_snap() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = SliverAppBar(floating: true, snap: true);
  print(s.snap);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_app_bar_expanded_height() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = SliverAppBar(expandedHeight: 200.0);
  print(s.expandedHeight);
}
"#
        ),
        vec!["200.0"]
    );
}

#[test]
fn sliver_app_bar_flexible_space() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = SliverAppBar(flexibleSpace: FlexibleSpaceBar(title: const Text('Flex')));
  print(s.flexibleSpace is FlexibleSpaceBar);
}
"#
        ),
        vec!["true"]
    );
}
