// vybe-test: dart/dart_developer_inspect_log/developer_log_basic
// origin: languages/dart/tests/dart/test_dart_developer_inspect_log.rs

import 'dart:developer';

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

void __vybeMain() {
  // Log message doesn't go to stdout by default, it goes to VM Service
  log('This is a test log');
  __p('log_called');
}

void main() {
  __vybeMain();
  __check('log_called');
}
