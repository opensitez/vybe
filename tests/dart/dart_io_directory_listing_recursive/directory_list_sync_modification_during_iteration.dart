// vybe-test: dart/dart_io_directory_listing_recursive/directory_list_sync_modification_during_iteration
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
  final dir = Directory.systemTemp.createTempSync('list_mod_');
  File('${dir.path}/1.txt').createSync();
  
  int count = 0;
  // listSync returns a list, so modification doesn't affect the already-returned list
  final items = dir.listSync();
  File('${dir.path}/2.txt').createSync();
  print(items.length);
  dir.deleteSync(recursive: true);
}

void main() {
  __vybeMain();
  __check('1');
}
