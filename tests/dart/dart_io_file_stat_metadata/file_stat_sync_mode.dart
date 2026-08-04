// vybe-test: dart/dart_io_file_stat_metadata/file_stat_sync_mode
// origin: languages/dart/tests/dart/test_dart_io_file_stat_metadata.rs

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
  final file = File('test_stat_mode.txt');
  file.writeAsStringSync('mode');
  final stat = file.statSync();
  // We can't guarantee the exact mode across OSes, just that it's > 0
  __p(stat.mode > 0);
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('true');
}
