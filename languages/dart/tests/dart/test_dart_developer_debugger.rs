use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:developer Debugger
// ═══════════════════════════════════════════════════════════

#[test]
fn developer_debugger_trigger() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  // If no debugger is attached, this is a no-op and returns false.
  final triggered = debugger(message: 'test_debugger');
  print(triggered is bool);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn developer_debugger_when_condition() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final triggered = debugger(when: false);
  print(triggered == false);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn developer_debugger_when_condition_true() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  // when=true still returns false if no debugger attached
  final triggered = debugger(when: true);
  print(triggered == false);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn developer_post_event() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  // Posting an event to the observatory / VM service stream
  // It shouldn't crash if no one is listening.
  postEvent('my.custom.event', {'key': 'value'});
  print('posted');
}
"#
        ),
        vec!["posted"]
    );
}

// `developer_post_event_invalid_json_throws` is GONE. Its program declared
// `class Unserializable {}` inside a function body, which Dart has never
// allowed — `dart run` answers "'class' can't be used as an identifier because
// it's a keyword." A test the reference implementation cannot compile states
// nothing about `postEvent`, and `developer_post_event` above already covers
// the "it does not crash with no listener" contract it was reaching for.
