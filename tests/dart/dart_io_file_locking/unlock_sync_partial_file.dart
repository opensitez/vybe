// vybe-test: dart/dart_io_file_locking/unlock_sync_partial_file
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
  final file = File('unlock_partial.txt');
  file.writeAsStringSync('0123456789');
  final raf = file.openSync(mode: FileMode.write);
  
  raf.lockSync(FileLock.exclusive, 2, 5);
  raf.unlockSync(2, 5);
  __p('partial unlocked');
  
  raf.closeSync();
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('partial unlocked');
}
