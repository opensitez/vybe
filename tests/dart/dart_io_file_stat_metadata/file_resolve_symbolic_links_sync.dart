// vybe-test: dart/dart_io_file_stat_metadata/file_resolve_symbolic_links_sync
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
  final file = File('test_sym_resolve.txt');
  file.writeAsStringSync('data');
  final resolved = file.resolveSymbolicLinksSync();
  __p(resolved.isNotEmpty);
  file.deleteSync();
}

void main() {
  __vybeMain();
  __check('true');
}
