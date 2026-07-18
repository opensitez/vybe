use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material BottomNavigationBar
// ═══════════════════════════════════════════════════════════

#[test]
fn bottom_navigation_bar_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomNavigationBar(
    items: const [
      BottomNavigationBarItem(icon: Icon(Icons.ac_unit), label: 'A'),
      BottomNavigationBarItem(icon: Icon(Icons.access_alarm), label: 'B'),
    ],
  );
  print(b is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bottom_navigation_bar_items_count() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomNavigationBar(
    items: const [
      BottomNavigationBarItem(icon: Icon(Icons.ac_unit), label: 'A'),
      BottomNavigationBarItem(icon: Icon(Icons.access_alarm), label: 'B'),
      BottomNavigationBarItem(icon: Icon(Icons.add), label: 'C'),
    ],
  );
  print(b.items.length);
}
"#
        ),
        vec!["3"]
    );
}

#[test]
fn bottom_navigation_bar_current_index() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomNavigationBar(
    currentIndex: 1,
    items: const [
      BottomNavigationBarItem(icon: Icon(Icons.ac_unit), label: 'A'),
      BottomNavigationBarItem(icon: Icon(Icons.access_alarm), label: 'B'),
    ],
  );
  print(b.currentIndex);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn bottom_navigation_bar_elevation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomNavigationBar(
    elevation: 8.0,
    items: const [
      BottomNavigationBarItem(icon: Icon(Icons.ac_unit), label: 'A'),
      BottomNavigationBarItem(icon: Icon(Icons.access_alarm), label: 'B'),
    ],
  );
  print(b.elevation);
}
"#
        ),
        vec!["8.0"]
    );
}

#[test]
fn bottom_navigation_bar_type() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomNavigationBar(
    type: BottomNavigationBarType.fixed,
    items: const [
      BottomNavigationBarItem(icon: Icon(Icons.ac_unit), label: 'A'),
      BottomNavigationBarItem(icon: Icon(Icons.access_alarm), label: 'B'),
    ],
  );
  print(b.type == BottomNavigationBarType.fixed);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bottom_navigation_bar_background_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final b = BottomNavigationBar(
    backgroundColor: const Color(0xFF112233),
    items: const [
      BottomNavigationBarItem(icon: Icon(Icons.ac_unit), label: 'A'),
      BottomNavigationBarItem(icon: Icon(Icons.access_alarm), label: 'B'),
    ],
  );
  print(b.backgroundColor?.value == 0xFF112233);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn bottom_navigation_bar_item_properties() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final item = BottomNavigationBarItem(
    icon: const Icon(Icons.home),
    activeIcon: const Icon(Icons.home_filled),
    label: 'Home',
    backgroundColor: const Color(0xFF00FF00),
  );
  print('${item.label}:${item.backgroundColor?.value == 0xFF00FF00}');
}
"#
        ),
        vec!["Home:true"]
    );
}
