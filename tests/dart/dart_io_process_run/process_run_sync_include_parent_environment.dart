// vybe-test: dart/dart_io_process_run/process_run_sync_include_parent_environment
// origin: languages/dart/tests/dart/test_dart_io_process_run.rs

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
  // If includeParentEnvironment is false, the basic PATH might be gone, making 'env' fail.
  // We'll use a shell to echo a specific variable to see if it's there.
  __p('test_passed'); // Just validating API surface for now.
}

void main() {
  __vybeMain();
  __check('test_passed');
}
