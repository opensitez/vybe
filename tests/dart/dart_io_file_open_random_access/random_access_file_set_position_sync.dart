// vybe-test: dart/dart_io_file_open_random_access/random_access_file_set_position_sync
// origin: languages/dart/tests/dart/test_dart_io_file_open_random_access.rs

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
  final file = File('test_raf_seek.bin');
  file.writeAsBytesSync([10, 20, 30, 40]);
  final raf = file.openSync(mode: FileMode.read);
  raf.setPositionSync(2);
  __p(raf.readByteSync());
  raf.closeSync();
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('30');
}
