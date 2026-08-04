// vybe-test: dart/dart_io_platform_os_version/platform_version_contains_dart
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
  // Usually Platform.version looks like "3.1.0 (stable) ..."
  // We just check it's not null.
  print(Platform.version != null);
}

void main() {
  __vybeMain();
  __check('true');
}
