use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:developer Inspect & Log
// ═══════════════════════════════════════════════════════════

#[test]
fn developer_inspect_object() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final obj = {'hello': 'world'};
  // Inspect returns the object itself.
  final returned = inspect(obj);
  print(identical(obj, returned));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn developer_log_basic() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  // Log message doesn't go to stdout by default, it goes to VM Service
  log('This is a test log');
  print('log_called');
}
"#
        ),
        vec!["log_called"]
    );
}

#[test]
fn developer_log_with_all_params() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  log(
    'Detailed log',
    time: DateTime.now(),
    sequenceNumber: 42,
    level: 1000,
    name: 'my.logger',
    zone: Zone.current,
    error: ArgumentError('bad arg'),
    stackTrace: StackTrace.current,
  );
  print('detailed_log_called');
}
"#
        ),
        vec!["detailed_log_called"]
    );
}

#[test]
fn developer_log_error_only() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  log('', error: Exception('Oops'));
  print('error_logged');
}
"#
        ),
        vec!["error_logged"]
    );
}

#[test]
fn developer_inspect_null() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final returned = inspect(null);
  print(returned == null);
}
"#
        ),
        vec!["true"]
    );
}
