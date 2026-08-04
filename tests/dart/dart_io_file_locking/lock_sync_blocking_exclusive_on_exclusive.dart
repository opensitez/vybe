// vybe-test: dart/dart_io_file_locking/lock_sync_blocking_exclusive_on_exclusive
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
  final file = File('lock_blk_ex_ex.txt');
  file.writeAsStringSync('data');
  final raf1 = file.openSync(mode: FileMode.write);
  final raf2 = file.openSync(mode: FileMode.write);
  
  raf1.lockSync(FileLock.exclusive);
  
  try {
    raf2.lockSync(FileLock.blockingExclusive);
    // Since we're in the same isolate/process in testing, behaviour might vary
    // Typically it would block or throw if we try non-blocking
    print('blocking...');
  } catch (e) {
    print('error');
  } finally {
    raf1.closeSync();
    raf2.closeSync();
    file.deleteSync();
  }
}

void main() {
  __vybeMain();
  __check('blocking...');
}
