// vybe-test: dart/dart_io_directory_symlinks/link_create_sync_to_directory
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
  final dir = Directory('test_link_dir');
  dir.createSync();
  final link = Link('test_link_dir.lnk');
  link.createSync(dir.path);
  __p(link.existsSync());
  link.deleteSync();
  dir.deleteSync();
}

void main() {
  __vybeMain();
  __check('true');
}
