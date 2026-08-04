// vybe-test: dart/dart_io_directory_symlinks/link_rename_sync_changes_path
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
  final link = Link('link1.lnk');
  link.createSync('tgt');
  final renamed = link.renameSync('link2.lnk');
  __p(renamed.path);
  renamed.deleteSync();
}

void main() {
  __vybeMain();
  __check('link2.lnk');
}
