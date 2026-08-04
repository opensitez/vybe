// vybe-test: dart/dart_io_process_start_streams/process_stdin_add_error
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
  final process = await Process.start('echo', ['done']);
  await process.exitCode; // wait for it to die
  try {
    process.stdin.addError(Exception('error'));
    // Since stdin is closed/dead, addError might throw StateError or ignore
    __p('added');
  } catch(e) {
    __p('StateError thrown');
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('added');
}
