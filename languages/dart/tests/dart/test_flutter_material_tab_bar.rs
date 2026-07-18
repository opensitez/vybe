use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material TabBar
// ═══════════════════════════════════════════════════════════

#[test]
fn tab_bar_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final t = TabBar(tabs: const [Tab(text: 'A')]);
  print(t is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn tab_bar_tabs() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final t = TabBar(
    tabs: const [
      Tab(text: 'A'),
      Tab(text: 'B'),
    ],
  );
  print(t.tabs.length);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn tab_bar_is_scrollable() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final t = TabBar(
    isScrollable: true,
    tabs: const [Tab(text: 'A')],
  );
  print(t.isScrollable);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn tab_bar_indicator_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final t = TabBar(
    indicatorColor: const Color(0xFF00FF00),
    tabs: const [Tab(text: 'A')],
  );
  print(t.indicatorColor?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn tab_bar_label_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final t = TabBar(
    labelColor: const Color(0xFF123456),
    tabs: const [Tab(text: 'A')],
  );
  print(t.labelColor?.value == 0xFF123456);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn tab_bar_unselected_label_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final t = TabBar(
    unselectedLabelColor: const Color(0xFF654321),
    tabs: const [Tab(text: 'A')],
  );
  print(t.unselectedLabelColor?.value == 0xFF654321);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn tab_bar_indicator_weight() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final t = TabBar(
    indicatorWeight: 4.0,
    tabs: const [Tab(text: 'A')],
  );
  print(t.indicatorWeight);
}
"#
        ),
        vec!["4.0"]
    );
}
