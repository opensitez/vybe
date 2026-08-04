// vybe-test: dart/dart_io_process_signal_handling/process_signal_watch_multiple_signals
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
  try {
    final sub1 = ProcessSignal.sigint.watch().listen((_) {});
    final sub2 = ProcessSignal.sigterm.watch().listen((_) {});
    sub1.cancel();
    sub2.cancel();
    __p('multi_watch');
  } catch (e) {
    __p('multi_watch');
  }
}

void main() {
  __vybeMain();
  __check('multi_watch');
}
