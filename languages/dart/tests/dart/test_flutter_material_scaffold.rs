use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Scaffold
// ═══════════════════════════════════════════════════════════

#[test]
fn scaffold_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Scaffold();
  print(s is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scaffold_app_bar() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Scaffold(appBar: AppBar());
  print(s.appBar != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scaffold_body() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Scaffold(body: const Placeholder());
  print(s.body is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scaffold_floating_action_button() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Scaffold(
    floatingActionButton: FloatingActionButton(onPressed: () {}),
  );
  print(s.floatingActionButton is FloatingActionButton);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scaffold_drawer() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Scaffold(drawer: Drawer());
  print(s.drawer is Drawer);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scaffold_bottom_navigation_bar() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Scaffold(
    bottomNavigationBar: BottomNavigationBar(
      items: const [
        BottomNavigationBarItem(icon: Icon(Icons.home), label: 'Home'),
        BottomNavigationBarItem(icon: Icon(Icons.settings), label: 'Settings'),
      ],
    ),
  );
  print(s.bottomNavigationBar is BottomNavigationBar);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scaffold_background_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Scaffold(backgroundColor: const Color(0xFF123456));
  print(s.backgroundColor?.value == 0xFF123456);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scaffold_of() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  print(Scaffold.maybeOf != null);
}
"#
        ),
        vec!["true"]
    );
}
