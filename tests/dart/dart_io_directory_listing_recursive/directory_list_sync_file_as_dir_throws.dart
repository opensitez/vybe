// vybe-test: dart/dart_io_directory_listing_recursive/directory_list_sync_file_as_dir_throws
// origin: languages/dart/tests/dart/test_dart_io_directory_listing_recursive.rs

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
  final file = File('test_list_file.txt');
  file.writeAsStringSync('data');
  final dir = Directory(file.path);
  try {
    dir.listSync();
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
