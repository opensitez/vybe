use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Chip
// ═══════════════════════════════════════════════════════════

#[test]
fn chip_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const c = Chip(label: Text('A'));
  print(c is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn chip_label() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const c = Chip(label: Text('Label'));
  print(c.label is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn chip_avatar() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const c = Chip(
    avatar: CircleAvatar(),
    label: Text('A'),
  );
  print(c.avatar is CircleAvatar);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn chip_delete_icon() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const c = Chip(
    label: Text('A'),
    deleteIcon: Icon(Icons.cancel),
    onDeleted: _dummy,
  );
  print(c.deleteIcon is Icon);
}
void _dummy() {}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn action_chip_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final ac = ActionChip(
    label: const Text('Action'),
    onPressed: () {},
  );
  print(ac is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn filter_chip_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final fc = FilterChip(
    label: const Text('Filter'),
    selected: true,
    onSelected: (bool value) {},
  );
  print(fc.selected);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn choice_chip_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final cc = ChoiceChip(
    label: const Text('Choice'),
    selected: false,
    onSelected: (bool value) {},
  );
  print(cc.selected);
}
"#
        ),
        vec!["false"]
    );
}
