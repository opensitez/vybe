// vybe-test: dart/dart_io_link_resolution/resolve_symbolic_links_sync_cyclic_throws
// origin: languages/dart/tests/dart/test_dart_io_link_resolution.rs

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
  final l1 = Link('c1.lnk');
  final l2 = Link('c2.lnk');
  l1.createSync('c2.lnk');
  l2.createSync('c1.lnk');
  try {
    l1.resolveSymbolicLinksSync();
  } on FileSystemException {
    __p('FileSystemException thrown');
  } finally {
    l1.deleteSync();
    l2.deleteSync();
  }
}

void main() {
  __vybeMain();
  __check('FileSystemException thrown');
}
