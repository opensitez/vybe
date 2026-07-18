use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material DropdownButton
// ═══════════════════════════════════════════════════════════

#[test]
fn dropdown_button_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final db = DropdownButton<String>(
    items: const [],
    onChanged: (String? newValue) {},
  );
  print(db is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dropdown_button_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final db = DropdownButton<int>(
    value: 5,
    items: const [],
    onChanged: (int? newValue) {},
  );
  print(db.value);
}
"#
        ),
        vec!["5"]
    );
}

#[test]
fn dropdown_button_items() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final db = DropdownButton<String>(
    items: const [
      DropdownMenuItem(value: 'A', child: Text('A')),
      DropdownMenuItem(value: 'B', child: Text('B')),
    ],
    onChanged: (String? newValue) {},
  );
  print(db.items?.length);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn dropdown_button_icon() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final db = DropdownButton<String>(
    icon: const Icon(Icons.arrow_downward),
    items: const [],
    onChanged: (String? newValue) {},
  );
  print(db.icon is Icon);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dropdown_button_is_expanded() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final db = DropdownButton<String>(
    isExpanded: true,
    items: const [],
    onChanged: (String? newValue) {},
  );
  print(db.isExpanded);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dropdown_menu_item_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dmi = DropdownMenuItem<int>(
    value: 42,
    child: const Text('42'),
  );
  print('${dmi.value}:${dmi.child is Text}');
}
"#
        ),
        vec!["42:true"]
    );
}
