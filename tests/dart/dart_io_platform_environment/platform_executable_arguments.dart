// vybe-test: dart/dart_io_platform_environment/platform_executable_arguments
// origin: languages/dart/tests/dart/test_dart_io_platform_environment.rs

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
  final args = Platform.executableArguments;
  __p(args is List<String>);
}

void main() {
  __vybeMain();
  __check('true');
}
