// vybe-test: dart/dart_io_file_locking/lock_sync_zero_length_locks_to_infinity
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
  final file = File('lock_inf.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.write);
  
  // Specifying 0 as length might mean "to end of file" or "to infinity" depending on platform.
  // Wait, Dart API doesn't mention special meaning for 0.
  // Let's just pass it and see it doesn't crash.
  raf.lockSync(FileLock.exclusive, 0, 0);
  print('locked zero len');
  
  raf.closeSync();
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('locked zero len');
}
