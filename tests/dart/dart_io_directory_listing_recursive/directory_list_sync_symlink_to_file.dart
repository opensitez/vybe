// vybe-test: dart/dart_io_directory_listing_recursive/directory_list_sync_symlink_to_file
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
  final dir = Directory.systemTemp.createTempSync('list_link_file_');
  final file = File('${dir.path}/f1.txt');
  file.createSync();
  Link('${dir.path}/l1').createSync(file.path);
  
  final items = dir.listSync();
  __p(items.length);
  dir.deleteSync(recursive: true);
}

void main() {
  __vybeMain();
  __check('2');
}
