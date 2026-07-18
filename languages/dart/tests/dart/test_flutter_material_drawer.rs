use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Drawer
// ═══════════════════════════════════════════════════════════

#[test]
fn drawer_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final d = Drawer();
  print(d is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn drawer_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final d = Drawer(child: const Placeholder());
  print(d.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn drawer_elevation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final d = Drawer(elevation: 10.0);
  print(d.elevation);
}
"#
        ),
        vec!["10.0"]
    );
}

#[test]
fn drawer_semantic_label() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final d = Drawer(semanticLabel: 'MyDrawer');
  print(d.semanticLabel);
}
"#
        ),
        vec!["MyDrawer"]
    );
}

#[test]
fn drawer_header_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dh = DrawerHeader(child: const SizedBox());
  print(dh is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn drawer_header_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dh = DrawerHeader(child: const Text('Header'));
  print(dh.child is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn drawer_header_margin() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dh = DrawerHeader(
    margin: const EdgeInsets.all(5.0),
    child: const SizedBox(),
  );
  print((dh.margin as EdgeInsets).top);
}
"#
        ),
        vec!["5.0"]
    );
}
