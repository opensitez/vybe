// vybe-test: dart/dart_isolate_capabilities_pause_kill/capability_equality
// origin: languages/dart/tests/dart/test_dart_isolate_capabilities_pause_kill.rs

import 'dart:isolate';

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
  final cap1 = Capability();
  final cap2 = Capability();
  __p(cap1 != cap2);
  __p(cap1 == cap1);
}

void main() {
  __vybeMain();
  __check('true\ntrue');
}
