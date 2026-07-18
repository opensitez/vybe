use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material TimePicker
// ═══════════════════════════════════════════════════════════

#[test]
fn time_picker_dialog_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tpd = TimePickerDialog(
    initialTime: const TimeOfDay(hour: 10, minute: 30),
  );
  print(tpd is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn time_picker_dialog_initial_time() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tpd = TimePickerDialog(
    initialTime: const TimeOfDay(hour: 10, minute: 30),
  );
  print('${tpd.initialTime.hour}:${tpd.initialTime.minute}');
}
"#
        ),
        vec!["10:30"]
    );
}

#[test]
fn time_picker_dialog_help_text() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tpd = TimePickerDialog(
    initialTime: const TimeOfDay(hour: 10, minute: 30),
    helpText: 'Select Time',
  );
  print(tpd.helpText);
}
"#
        ),
        vec!["Select Time"]
    );
}

#[test]
fn time_picker_dialog_cancel_text() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tpd = TimePickerDialog(
    initialTime: const TimeOfDay(hour: 10, minute: 30),
    cancelText: 'Dismiss',
  );
  print(tpd.cancelText);
}
"#
        ),
        vec!["Dismiss"]
    );
}

#[test]
fn time_picker_dialog_confirm_text() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final tpd = TimePickerDialog(
    initialTime: const TimeOfDay(hour: 10, minute: 30),
    confirmText: 'Done',
  );
  print(tpd.confirmText);
}
"#
        ),
        vec!["Done"]
    );
}
