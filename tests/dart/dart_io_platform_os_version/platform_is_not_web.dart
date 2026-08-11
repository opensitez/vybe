// vybe-test: dart/dart_io_platform_os_version/platform_is_not_web
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
  // If dart:io is available, we are definitely not on the web.
  // We can't directly check isWeb from dart:io, but this confirms we compiled dart:io.
  __p('dart_io_loaded');
}

void main() {
  __vybeMain();
  __check('dart_io_loaded');
}
