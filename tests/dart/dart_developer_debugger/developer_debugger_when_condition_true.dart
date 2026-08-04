// vybe-test: dart/dart_developer_debugger/developer_debugger_when_condition_true
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
  // when=true still returns false if no debugger attached
  final triggered = debugger(when: true);
  __p(triggered == false);
}

void main() {
  __vybeMain();
  __check('true');
}
