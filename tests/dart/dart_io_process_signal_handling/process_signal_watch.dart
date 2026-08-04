// vybe-test: dart/dart_io_process_signal_handling/process_signal_watch
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
  // Can't easily trigger signals in this test, but we can verify the stream is returned
  try {
    final stream = ProcessSignal.sighup.watch();
    print(stream is Stream<ProcessSignal>);
  } catch (e) {
    // Windows might throw SignalException for unsupported signals
    print('SignalException');
  }
}

void main() {
  __vybeMain();
  __check('true');
}
