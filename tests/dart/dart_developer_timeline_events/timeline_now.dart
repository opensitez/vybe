// vybe-test: dart/dart_developer_timeline_events/timeline_now
// origin: languages/dart/tests/dart/test_dart_developer_timeline_events.rs

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
  final now = Timeline.now;
  // It returns microseconds since some epoch
  __p(now > 0);
}

void main() {
  __vybeMain();
  __check('true');
}
