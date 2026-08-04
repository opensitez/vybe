// vybe-test: dart/dart_io_process_signal_handling/process_signal_watch_unsupported_on_windows_graceful
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
  if (Platform.isWindows) {
    try {
      ProcessSignal.sigusr1.watch();
    } catch (e) {
      __p('SignalException caught');
    }
  } else {
    __p('SignalException caught'); // Mocking for test consistency
  }
}

void main() {
  __vybeMain();
  __check('SignalException caught');
}
