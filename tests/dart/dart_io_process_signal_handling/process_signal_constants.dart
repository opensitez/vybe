// vybe-test: dart/dart_io_process_signal_handling/process_signal_constants
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
  __p(ProcessSignal.sighup != null);
  __p(ProcessSignal.sigquit != null);
  __p(ProcessSignal.sigterm != null);
  __p(ProcessSignal.sigusr1 != null);
  __p(ProcessSignal.sigusr2 != null);
}

void main() {
  __vybeMain();
  __check('true\ntrue\ntrue\ntrue\ntrue');
}
