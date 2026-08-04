// vybe-test: dart/dart_io_process_start_streams/process_start_include_parent_environment_false
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
  try {
    // ls usually depends on PATH unless absolute path is provided.
    // If includeParentEnvironment is false, it might fail to find 'env' or 'ls' depending on platform.
    final process = await Process.start('env', [], includeParentEnvironment: false, environment: {'A': 'B'});
    __p(process.pid > 0);
  } catch (e) {
    __p('failed without path');
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('');
}
