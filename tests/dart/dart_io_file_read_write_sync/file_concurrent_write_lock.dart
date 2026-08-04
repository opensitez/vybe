// vybe-test: dart/dart_io_file_read_write_sync/file_concurrent_write_lock
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
  final file = File('test_write_lock.txt');
  file.writeAsStringSync('write lock');
  final raf = file.openSync(mode: FileMode.write);
  raf.lockSync(FileLock.exclusive);
  __p('Write locked successfully');
  raf.closeSync();
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('Write locked successfully');
}
