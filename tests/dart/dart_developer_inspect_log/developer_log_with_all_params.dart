// vybe-test: dart/dart_developer_inspect_log/developer_log_with_all_params
// origin: languages/dart/tests/dart/test_dart_developer_inspect_log.rs

import 'dart:async';
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
  log(
    'Detailed log',
    time: DateTime.now(),
    sequenceNumber: 42,
    level: 1000,
    name: 'my.logger',
    zone: Zone.current,
    error: ArgumentError('bad arg'),
    stackTrace: StackTrace.current,
  );
  __p('detailed_log_called');
}

void main() {
  __vybeMain();
  __check('detailed_log_called');
}
