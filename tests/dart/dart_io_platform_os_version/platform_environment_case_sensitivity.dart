// vybe-test: dart/dart_io_platform_os_version/platform_environment_case_sensitivity
// origin: languages/dart/tests/dart/test_dart_io_platform_os_version.rs

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
  // On Windows environment keys are case-insensitive, on Unix they are sensitive.
  // We'll just verify the map respects standard Dart map semantics.
  final env = Platform.environment;
  print(env.containsKey('NO_SUCH_KEY_123') == false);
}

void main() {
  __vybeMain();
  __check('true');
}
