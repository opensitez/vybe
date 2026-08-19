// vybe-test: dart/dart_developer_debugger/developer_debugger_trigger
// origin: languages/dart/tests/dart/test_dart_developer_debugger.rs

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
  // If no debugger is attached, this is a no-op and returns false.
  final triggered = debugger(message: 'test_debugger');
  __p(triggered is bool);
}

void main() {
  __vybeMain();
  __check('true');
}
