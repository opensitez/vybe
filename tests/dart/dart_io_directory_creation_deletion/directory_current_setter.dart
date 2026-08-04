// vybe-test: dart/dart_io_directory_creation_deletion/directory_current_setter
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
  final original = Directory.current;
  final temp = Directory.systemTemp.createTempSync();
  Directory.current = temp;
  __p(Directory.current.path == temp.path);
  Directory.current = original;
  temp.deleteSync();
}

void main() {
  __vybeMain();
  __check('true');
}
