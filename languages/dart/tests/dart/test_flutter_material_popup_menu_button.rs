use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material PopupMenuButton
// ═══════════════════════════════════════════════════════════

#[test]
fn popup_menu_button_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final pmb = PopupMenuButton<String>(
    itemBuilder: (BuildContext context) => [],
  );
  print(pmb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn popup_menu_button_item_builder() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final pmb = PopupMenuButton<int>(
    itemBuilder: (BuildContext context) => <PopupMenuEntry<int>>[
      const PopupMenuItem<int>(value: 1, child: Text('1')),
    ],
  );
  print(pmb.itemBuilder != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn popup_menu_button_initial_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final pmb = PopupMenuButton<String>(
    initialValue: 'A',
    itemBuilder: (BuildContext context) => [],
  );
  print(pmb.initialValue);
}
"#
        ),
        vec!["A"]
    );
}

#[test]
fn popup_menu_button_on_selected() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final pmb = PopupMenuButton<int>(
    onSelected: (int value) {},
    itemBuilder: (BuildContext context) => [],
  );
  print(pmb.onSelected != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn popup_menu_button_icon() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final pmb = PopupMenuButton<String>(
    icon: const Icon(Icons.more_vert),
    itemBuilder: (BuildContext context) => [],
  );
  print(pmb.icon is Icon);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn popup_menu_item_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final pmi = PopupMenuItem<double>(
    value: 3.14,
    child: const Text('Pi'),
  );
  print('${pmi.value}:${pmi.child is Text}');
}
"#
        ),
        vec!["3.14:true"]
    );
}
