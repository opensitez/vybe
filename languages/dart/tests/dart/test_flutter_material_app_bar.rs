use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material AppBar
// ═══════════════════════════════════════════════════════════

#[test]
fn app_bar_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final a = AppBar();
  print(a is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn app_bar_title() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final a = AppBar(title: const Text('Title'));
  print(a.title is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn app_bar_actions() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final a = AppBar(actions: [const IconButton(icon: Icon(Icons.add), onPressed: null)]);
  print(a.actions?.length);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn app_bar_leading() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final a = AppBar(leading: const BackButton());
  print(a.leading is BackButton);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn app_bar_bottom() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final a = AppBar(
    bottom: const TabBar(tabs: [Tab(text: 'A')]),
  );
  print(a.bottom is TabBar);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn app_bar_elevation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final a = AppBar(elevation: 4.0);
  print(a.elevation);
}
"#
        ),
        vec!["4.0"]
    );
}

#[test]
fn app_bar_background_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final a = AppBar(backgroundColor: const Color(0xFF112233));
  print(a.backgroundColor?.value == 0xFF112233);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn app_bar_center_title() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final a = AppBar(centerTitle: true);
  print(a.centerTitle);
}
"#
        ),
        vec!["true"]
    );
}
