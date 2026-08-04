// vybe-test: dart/dart_io_directory_listing_recursive/directory_list_sync_cyclic_links_error
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
  final dir = Directory.systemTemp.createTempSync('list_cyclic_');
  // Create a link pointing to its own parent
  Link('${dir.path}/cycle').createSync(dir.path);
  try {
    dir.listSync(recursive: true, followLinks: true).length;
    __p('Did not throw'); // Dart throws FileSystemException for cyclic links
  } on FileSystemException {
    __p('FileSystemException thrown');
  } finally {
    dir.deleteSync(recursive: true);
  }
}

void main() {
  __vybeMain();
  __check('FileSystemException thrown');
}
