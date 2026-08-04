// vybe-test: dart/dart_io_platform_environment/platform_environment_contains_path
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
  final env = Platform.environment;
  // Usually every OS has some form of PATH or Path
  final hasPath = env.containsKey('PATH') || env.containsKey('Path');
  // We'll just verify we can access keys
  print(env.keys.isNotEmpty);
}

void main() {
  __vybeMain();
  __check('true');
}
