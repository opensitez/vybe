// vybe-test: dart/dart_isolate_errors_exit/isolate_current_set_errors_fatal
// origin: languages/dart/tests/dart/test_dart_isolate_errors_exit.rs

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

import 'dart:isolate';
void __vybeMain() {
  Isolate.current.setErrorsFatal(false);
  __p('set');
}

void main() {
  __vybeMain();
  __check('set');
}
