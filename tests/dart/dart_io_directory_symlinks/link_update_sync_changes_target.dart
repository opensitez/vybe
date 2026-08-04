// vybe-test: dart/dart_io_directory_symlinks/link_update_sync_changes_target
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
  final link = Link('update_link.lnk');
  link.createSync('tgt1.txt');
  link.updateSync('tgt2.txt');
  __p(link.targetSync());
  link.deleteSync();
}

void main() {
  __vybeMain();
  __check('tgt2.txt');
}
