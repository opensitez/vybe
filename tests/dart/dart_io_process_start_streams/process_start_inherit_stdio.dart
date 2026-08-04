// vybe-test: dart/dart_io_process_start_streams/process_start_inherit_stdio
// origin: languages/dart/tests/dart/test_dart_io_process_start_streams.rs

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
void __vybeMain() async {
  final process = await Process.start('echo', ['inherit'], mode: ProcessStartMode.inheritStdio);
  // stdout/stderr are null when inherited
  __p(process.stdout == null);
}

Future<void> main() async {
  await __vybeMain();
  __check('test_inherit');
}
