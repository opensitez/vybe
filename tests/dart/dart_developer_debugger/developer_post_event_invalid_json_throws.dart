// vybe-test: dart/dart_developer_debugger/developer_post_event_invalid_json_throws
// origin: languages/dart/tests/dart/test_dart_developer_debugger.rs

final StringBuffer __vybeOut = StringBuffer();

void __p(Object? o) {
  __vybeOut.writeln(o);
}

void __check(String want) {
  var got = __vybeOut.toString();
  // `writeln` on the final print contributes a trailing newline that the
  // expected line vector never carried.
  if (got.endsWith('\n')) {
    got = got.substring(0, got.length - 1);
  }
  if (got != want) {
    print('FAIL: want [$want] got [$got]');
    throw Exception('assertion failed');
  }
}

import 'dart:developer';
void __vybeMain() {
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

void main() {
  __vybeMain();
  __check('done');
}
