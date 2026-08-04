// vybe-test: dart/dart_developer_inspect_log/developer_log_error_only
// origin: languages/dart/tests/dart/test_dart_developer_inspect_log.rs

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
  log('', error: Exception('Oops'));
  __p('error_logged');
}

void main() {
  __vybeMain();
  __check('error_logged');
}
