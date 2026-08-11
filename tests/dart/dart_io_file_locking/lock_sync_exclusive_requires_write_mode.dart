// vybe-test: dart/dart_io_file_locking/lock_sync_exclusive_requires_write_mode
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
  final file = File('lock_ex_read_mode.txt');
  file.writeAsStringSync('data');
  // Opened in read-only mode
  final raf = file.openSync(mode: FileMode.read);
  
  try {
    // Attempting exclusive lock on read-only descriptor throws
    raf.lockSync(FileLock.exclusive);
    __p('locked'); // shouldn't happen on strict OS, but may on some.
  } on FileSystemException {
    __p('FileSystemException thrown');
  } finally {
    raf.closeSync();
    file.deleteSync();
  }
}

void main() {
  __vybeMain();
  __check('FileSystemException thrown');
}
