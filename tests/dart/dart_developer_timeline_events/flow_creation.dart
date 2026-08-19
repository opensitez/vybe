// vybe-test: dart/dart_developer_timeline_events/flow_creation
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
  final flow = Flow.begin();
  __p(flow.id > 0);
  Flow.step(flow.id);
  Flow.end(flow.id);
}

void main() {
  __vybeMain();
  __check('true');
}
