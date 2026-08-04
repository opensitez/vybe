// vybe-test: dart/dart_io_file_read_write_sync/file_read_as_bytes_sync_large
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
  final file = File('test_bytes_sync.bin');
  file.writeAsBytesSync([0, 255, 128, 64]);
  final bytes = file.readAsBytesSync();
  __p('${bytes.length}:${bytes[1]}');
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('4:255');
}
