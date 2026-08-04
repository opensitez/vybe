// vybe-test: dart/dart_io_directory_creation_deletion/directory_create_invalid_characters_throws
// origin: languages/dart/tests/dart/test_dart_io_directory_creation_deletion.rs

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
  final dir = Directory('invalid\x00dir');
  try {
    dir.createSync();
  } on FileSystemException {
    __p('FileSystemException thrown');
  } catch (e) {
    __p('ArgumentError thrown'); // Depending on platform, may throw ArgumentError
  }
}

void main() {
  __vybeMain();
  __check('ArgumentError thrown');
}
