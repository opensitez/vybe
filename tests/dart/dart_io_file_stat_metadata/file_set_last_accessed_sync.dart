// vybe-test: dart/dart_io_file_stat_metadata/file_set_last_accessed_sync
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
  final file = File('test_set_last_acc.txt');
  file.writeAsStringSync('test');
  final target = DateTime(2031, 2, 2);
  file.setLastAccessedSync(target);
  final stat = file.statSync();
  __p('${stat.accessed.year}:${stat.accessed.month}:${stat.accessed.day}');
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('2031:2:2');
}
