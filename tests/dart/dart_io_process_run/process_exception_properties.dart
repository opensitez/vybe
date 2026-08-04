// vybe-test: dart/dart_io_process_run/process_exception_properties
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
  try {
    Process.runSync('does_not_exist_proc', ['arg1']);
  } on ProcessException catch (e) {
    __p(e.executable == 'does_not_exist_proc');
    __p(e.arguments[0] == 'arg1');
    __p(e.message.isNotEmpty);
    __p(e.errorCode != 0);
  }
}

void main() {
  __vybeMain();
  __check('true\ntrue\ntrue\ntrue');
}
