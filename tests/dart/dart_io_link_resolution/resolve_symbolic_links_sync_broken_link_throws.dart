// vybe-test: dart/dart_io_link_resolution/resolve_symbolic_links_sync_broken_link_throws
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
  final link = Link('broken.lnk');
  link.createSync('non_existent.txt');
  try {
    link.resolveSymbolicLinksSync();
  } on FileSystemException {
    __p('FileSystemException thrown');
  } finally {
    link.deleteSync();
  }
}

void main() {
  __vybeMain();
  __check('FileSystemException thrown');
}
