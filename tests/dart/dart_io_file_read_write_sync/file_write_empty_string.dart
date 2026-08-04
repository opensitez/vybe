// vybe-test: dart/dart_io_file_read_write_sync/file_write_empty_string
// origin: languages/dart/tests/dart/test_dart_io_file_read_write_sync.rs

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
  final file = File('test_empty_sync.txt');
  file.writeAsStringSync('');
  __p(file.lengthSync());
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('0');
}
