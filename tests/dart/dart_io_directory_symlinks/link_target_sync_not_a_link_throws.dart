// vybe-test: dart/dart_io_directory_symlinks/link_target_sync_not_a_link_throws
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
  final file = File('not_a_link.txt');
  file.createSync();
  final link = Link(file.path); // Point Link object to an actual File
  try {
    link.targetSync();
  } on FileSystemException {
    __p('FileSystemException thrown');
  } finally {
    file.deleteSync();
  }
}

void main() {
  __vybeMain();
  __check('FileSystemException thrown');
}
