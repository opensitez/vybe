// vybe-test: dart/dart_developer_debugger/developer_post_event
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
  // Posting an event to the observatory / VM service stream
  // It shouldn't crash if no one is listening.
  postEvent('my.custom.event', {'key': 'value'});
  print('posted');
}

void main() {
  __vybeMain();
  __check('posted');
}
