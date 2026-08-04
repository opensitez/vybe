// vybe-test: dart/dart_io_file_system_entity_watch/link_watch_returns_stream
// origin: languages/dart/tests/dart/test_dart_io_file_system_entity_watch.rs

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
  final link = Link('watch_link.lnk');
  link.createSync('dummy');
  final stream = link.watch();
  __p(stream is Stream<FileSystemEvent>);
  link.deleteSync();
}

void main() {
  __vybeMain();
  __check('true');
}
