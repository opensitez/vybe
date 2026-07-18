use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material ListTile
// ═══════════════════════════════════════════════════════════

#[test]
fn list_tile_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final lt = ListTile(title: const Text('Title'));
  print(lt is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_tile_title() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final lt = ListTile(title: const Text('Hello'));
  print(lt.title is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_tile_subtitle() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final lt = ListTile(subtitle: const Text('Sub'));
  print(lt.subtitle is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_tile_leading() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final lt = ListTile(leading: const Icon(Icons.person));
  print(lt.leading is Icon);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_tile_trailing() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final lt = ListTile(trailing: const Icon(Icons.arrow_forward));
  print(lt.trailing is Icon);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_tile_is_three_line() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final lt = ListTile(isThreeLine: true);
  print(lt.isThreeLine);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_tile_dense() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final lt = ListTile(dense: true);
  print(lt.dense);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_tile_on_tap() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final lt = ListTile(onTap: () {});
  print(lt.onTap != null);
}
"#
        ),
        vec!["true"]
    );
}
