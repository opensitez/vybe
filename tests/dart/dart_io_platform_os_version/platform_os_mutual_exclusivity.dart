// vybe-test: dart/dart_io_platform_os_version/platform_os_mutual_exclusivity
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
  int count = 0;
  if (Platform.isAndroid) count++;
  if (Platform.isFuchsia) count++;
  if (Platform.isIOS) count++;
  if (Platform.isLinux) count++;
  if (Platform.isMacOS) count++;
  if (Platform.isWindows) count++;
  // Web is not covered by dart:io (throws UnsupportedError if you try to import it on web)
  // Therefore, exact 1 OS should be true.
  __p(count == 1);
}

void main() {
  __vybeMain();
  __check('true');
}
