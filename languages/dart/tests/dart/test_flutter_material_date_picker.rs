use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material DatePicker
// ═══════════════════════════════════════════════════════════

#[test]
fn date_picker_dialog_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dpd = DatePickerDialog(
    initialDate: DateTime(2023, 1, 1),
    firstDate: DateTime(2000),
    lastDate: DateTime(2050),
  );
  print(dpd is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn date_picker_dialog_initial_date() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dpd = DatePickerDialog(
    initialDate: DateTime(2023, 1, 1),
    firstDate: DateTime(2000),
    lastDate: DateTime(2050),
  );
  print(dpd.initialDate.year);
}
"#
        ),
        vec!["2023"]
    );
}

#[test]
fn date_picker_dialog_first_date() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dpd = DatePickerDialog(
    initialDate: DateTime(2023, 1, 1),
    firstDate: DateTime(2000),
    lastDate: DateTime(2050),
  );
  print(dpd.firstDate.year);
}
"#
        ),
        vec!["2000"]
    );
}

#[test]
fn date_picker_dialog_last_date() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dpd = DatePickerDialog(
    initialDate: DateTime(2023, 1, 1),
    firstDate: DateTime(2000),
    lastDate: DateTime(2050),
  );
  print(dpd.lastDate.year);
}
"#
        ),
        vec!["2050"]
    );
}

#[test]
fn date_picker_dialog_help_text() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dpd = DatePickerDialog(
    initialDate: DateTime(2023, 1, 1),
    firstDate: DateTime(2000),
    lastDate: DateTime(2050),
    helpText: 'Select Date',
  );
  print(dpd.helpText);
}
"#
        ),
        vec!["Select Date"]
    );
}
