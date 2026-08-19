// vybe-test: dart/dart_developer_timeline_events/timeline_task_creation
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
  final task = TimelineTask();
  task.start('asyncTask');
  task.pass(); // pass down to nested
  task.finish();
  __p('task_done');
}

void main() {
  __vybeMain();
  __check('task_done');
}
