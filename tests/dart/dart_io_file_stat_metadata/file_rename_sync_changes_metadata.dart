// vybe-test: dart/dart_io_file_stat_metadata/file_rename_sync_changes_metadata
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
  final file = File('test_rename_meta.txt');
  file.writeAsStringSync('meta');
  final renamed = file.renameSync('test_renamed_meta.txt');
  __p(renamed.statSync().size);
  renamed.deleteSync();
}

void main() {
  __vybeMain();
  __check('4');
}
