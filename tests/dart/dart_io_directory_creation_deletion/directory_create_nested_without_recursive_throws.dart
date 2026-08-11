// vybe-test: dart/dart_io_directory_creation_deletion/directory_create_nested_without_recursive_throws
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
  final root = Directory('test_nested_1');
  if (root.existsSync()) {
    root.deleteSync(recursive: true);
  }
  final dir = Directory('test_nested_1/test_nested_2');
  try {
    dir.createSync();
  } on FileSystemException {
    __p('FileSystemException thrown');
  } finally {
    if (root.existsSync()) {
      root.deleteSync(recursive: true);
    }
  }
}

void main() {
  __vybeMain();
  __check('FileSystemException thrown');
}
