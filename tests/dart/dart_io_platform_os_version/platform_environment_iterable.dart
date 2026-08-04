// vybe-test: dart/dart_io_platform_os_version/platform_environment_iterable
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
  final env = Platform.environment;
  int count = 0;
  for (var k in env.keys) {
    count++;
  }
  __p(count == env.length);
}

void main() {
  __vybeMain();
  __check('true');
}
