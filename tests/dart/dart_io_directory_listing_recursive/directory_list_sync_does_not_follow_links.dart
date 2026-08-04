// vybe-test: dart/dart_io_directory_listing_recursive/directory_list_sync_does_not_follow_links
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
  final dir = Directory.systemTemp.createTempSync('list_no_links_');
  final targetDir = Directory.systemTemp.createTempSync('target_no_dir_');
  File('${targetDir.path}/f1.txt').createSync();
  Link('${dir.path}/l1').createSync(targetDir.path);
  
  final items = dir.listSync(recursive: true, followLinks: false);
  // It should see the link but NOT descend into it
  int links = items.whereType<Link>().length;
  int files = items.whereType<File>().length;
  __p('$links:$files');
  dir.deleteSync(recursive: true);
  targetDir.deleteSync(recursive: true);
}

void main() {
  __vybeMain();
  __check('1:0');
}
