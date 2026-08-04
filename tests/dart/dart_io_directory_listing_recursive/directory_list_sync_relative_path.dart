// vybe-test: dart/dart_io_directory_listing_recursive/directory_list_sync_relative_path
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
  final dir = Directory('relative_list_dir');
  dir.createSync();
  File('${dir.path}/f1.txt').createSync();
  final items = dir.listSync();
  // Ensure the returned paths are relative
  __p(!items[0].path.startsWith('/'));
  dir.deleteSync(recursive: true);
}

void main() {
  __vybeMain();
  __check('true');
}
