// vybe-test: dart/dart_io_directory_symlinks/link_create_cyclic_chain
// origin: languages/dart/tests/dart/test_dart_io_directory_symlinks.rs

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
  final link1 = Link('l1.lnk');
  final link2 = Link('l2.lnk');
  link1.createSync('l2.lnk');
  link2.createSync('l1.lnk');
  // targetSync just reads the link, it doesn't resolve it. So it won't infinite loop.
  __p(link1.targetSync() == 'l2.lnk');
  link1.deleteSync();
  link2.deleteSync();
}

void main() {
  __vybeMain();
  __check('true');
}
