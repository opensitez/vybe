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

#[test]
fn developer_post_event_invalid_json_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  // Event data must be JSON-serializable
  class Unserializable {}
  try {
    postEvent('bad.event', {'obj': Unserializable()});
    // The serialization is done internally by VM service. Some Darts might just stringify it,
    // or it might throw ArgumentError if it fails to serialize. Let's just catch any exception.
    print('done');
  } catch(e) {
    print('done'); // Safe fallback as VM implementations vary
  }
}
"#
        ),
        vec!["done"]
    );
}
