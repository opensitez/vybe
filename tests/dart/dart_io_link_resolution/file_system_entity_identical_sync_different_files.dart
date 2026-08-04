// vybe-test: dart/dart_io_link_resolution/file_system_entity_identical_sync_different_files
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
  final f1 = File('ident_f1.txt')..createSync();
  final f2 = File('ident_f2.txt')..createSync();
  __p(FileSystemEntity.identicalSync(f1.path, f2.path));
  f1.deleteSync();
  f2.deleteSync();
}

void main() {
  __vybeMain();
  __check('false');
}
