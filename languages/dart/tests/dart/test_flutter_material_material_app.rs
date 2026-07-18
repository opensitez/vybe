use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material MaterialApp
// ═══════════════════════════════════════════════════════════

#[test]
fn material_app_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ma = MaterialApp(home: const SizedBox());
  print(ma.title); // defaults to ''
}
"#
        ),
        vec![""]
    );
}

#[test]
fn material_app_title() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ma = MaterialApp(title: 'My App', home: const SizedBox());
  print(ma.title);
}
"#
        ),
        vec!["My App"]
    );
}

#[test]
fn material_app_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ma = MaterialApp(color: const Color(0xFF00FF00), home: const SizedBox());
  print(ma.color?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn material_app_theme() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ma = MaterialApp(
    theme: ThemeData.light(),
    home: const SizedBox(),
  );
  print(ma.theme != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn material_app_debug_show_checked_mode_banner() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ma = MaterialApp(
    debugShowCheckedModeBanner: false,
    home: const SizedBox(),
  );
  print(ma.debugShowCheckedModeBanner);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn material_app_initial_route() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ma = MaterialApp(
    initialRoute: '/home',
    routes: {'/home': (context) => const SizedBox()},
  );
  print(ma.initialRoute);
}
"#
        ),
        vec!["/home"]
    );
}

#[test]
fn material_app_is_stateful_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ma = MaterialApp(home: const SizedBox());
  print(ma is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}
