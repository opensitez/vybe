// vybe-test: dart/dart_io_link_resolution/file_system_entity_identical_sync_link_and_target
// origin: languages/dart/tests/dart/test_dart_io_link_resolution.rs

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
  final file = File('ident_tgt.txt')..createSync();
  final link = Link('ident_link.lnk')..createSync(file.path);
  __p(FileSystemEntity.identicalSync(file.path, link.path));
  link.deleteSync();
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('true');
}
