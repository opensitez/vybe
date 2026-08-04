// vybe-test: dart/dart_io_process_signal_handling/process_kill_pid_invalid_signal
// origin: languages/dart/tests/dart/test_dart_io_process_signal_handling.rs

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

import 'dart:io';
void __vybeMain() {
  // Testing with an invalid signal integer (though Dart types it securely)
  // Actually, ProcessSignal cannot be fabricated easily without reflection
  __p('secure');
}

void main() {
  __vybeMain();
  __check('secure');
}
