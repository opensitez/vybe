// vybe-test: dart/dart_io_platform_os_version/platform_os_string_matches_boolean
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
  final os = Platform.operatingSystem;
  if (os == 'android') __p(Platform.isAndroid);
  else if (os == 'fuchsia') __p(Platform.isFuchsia);
  else if (os == 'ios') __p(Platform.isIOS);
  else if (os == 'linux') __p(Platform.isLinux);
  else if (os == 'macos') __p(Platform.isMacOS);
  else if (os == 'windows') __p(Platform.isWindows);
  else __p('false');
}

void main() {
  __vybeMain();
  __check('true');
}
