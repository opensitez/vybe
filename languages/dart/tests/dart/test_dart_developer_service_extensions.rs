use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:developer Service Extensions
// ═══════════════════════════════════════════════════════════

#[test]
fn register_extension() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  registerExtension('ext.my.customMethod', (method, parameters) async {
    return ServiceExtensionResponse.result('{"success": true}');
  });
  print('extension_registered');
}
"#
        ),
        vec!["extension_registered"]
    );
}

#[test]
fn service_extension_response_result() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final response = ServiceExtensionResponse.result('{"key":"value"}');
  print(response.result);
}
"#
        ),
        vec!["{\"key\":\"value\"}"]
    );
}

#[test]
fn service_extension_response_error() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final response = ServiceExtensionResponse.error(ServiceExtensionResponse.invalidParams, 'bad args');
  print(response.errorCode);
  print(response.errorDetail);
}
"#
        ),
        vec!["-32602\nbad args"]
    );
}

#[test]
fn extension_stream_has_listener() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  // `extensionStreamHasListener` is a boolean property
  final hasListener = extensionStreamHasListener;
  print(hasListener is bool);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn post_event_multiple() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  postEvent('ext.test.event1', {'a': 1});
  postEvent('ext.test.event2', {'b': 2});
  print('posted_multiple');
}
"#
        ),
        vec!["posted_multiple"]
    );
}
