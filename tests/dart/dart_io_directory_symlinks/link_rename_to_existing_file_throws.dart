// vybe-test: dart/dart_io_directory_symlinks/link_rename_to_existing_file_throws
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
  final link = Link('link_to_rename.lnk');
  link.createSync('tgt');
  final file = File('existing_file_for_link.txt');
  file.createSync();
  try {
    link.renameSync(file.path);
  } on FileSystemException {
    __p('FileSystemException thrown');
  } finally {
    link.deleteSync();
    file.deleteSync();
  }
}

void main() {
  __vybeMain();
  __check('FileSystemException thrown');
}
