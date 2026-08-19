// vybe-test: dart/dart_developer_timeline_events/timeline_start_sync
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
  Timeline.startSync('myTask');
  Timeline.finishSync();
  __p('timeline_sync_done');
}

void main() {
  __vybeMain();
  __check('timeline_sync_done');
}
