// vybe-test: dart/dart_io_file_locking/lock_sync_concurrent_non_overlapping_ranges
// origin: languages/dart/tests/dart/test_dart_io_file_locking.rs

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
  final file = File('lock_non_overlap.txt');
  file.writeAsStringSync('0123456789');
  final raf1 = file.openSync(mode: FileMode.write);
  final raf2 = file.openSync(mode: FileMode.write);
  
  raf1.lockSync(FileLock.exclusive, 0, 4);
  raf2.lockSync(FileLock.exclusive, 5, 4); // Non-overlapping
  __p('locked both');
  
  raf1.closeSync();
  raf2.closeSync();
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('locked both');
}
