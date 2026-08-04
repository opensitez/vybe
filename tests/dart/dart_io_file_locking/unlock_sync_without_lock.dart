// vybe-test: dart/dart_io_file_locking/unlock_sync_without_lock
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
  final file = File('unlock_no_lock.txt');
  file.writeAsStringSync('data');
  final raf = file.openSync(mode: FileMode.read);
  
  // Depending on OS, unlocking an unlocked file might be a no-op or throw.
  try {
    raf.unlockSync();
    __p('unlocked no-op');
  } catch (e) {
    __p('error');
  }
  
  raf.closeSync();
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('unlocked no-op');
}
