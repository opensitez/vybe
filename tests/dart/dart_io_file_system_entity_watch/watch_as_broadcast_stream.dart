// vybe-test: dart/dart_io_file_system_entity_watch/watch_as_broadcast_stream
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
  final dir = Directory('watch_bcast');
  dir.createSync();
  final stream = dir.watch().asBroadcastStream();
  stream.listen((_) {});
  stream.listen((_) {});
  __p('success');
  dir.deleteSync();
}

void main() {
  __vybeMain();
  __check('success');
}
