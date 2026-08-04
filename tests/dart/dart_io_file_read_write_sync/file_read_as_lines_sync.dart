// vybe-test: dart/dart_io_file_read_write_sync/file_read_as_lines_sync
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
  final file = File('test_lines_sync.txt');
  file.writeAsStringSync('Line1\nLine2\r\nLine3');
  final lines = file.readAsLinesSync();
  __p('${lines.length}:${lines[0]}:${lines[2]}');
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('3:Line1:Line3');
}
