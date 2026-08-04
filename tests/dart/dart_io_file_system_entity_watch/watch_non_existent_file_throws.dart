// vybe-test: dart/dart_io_file_system_entity_watch/watch_non_existent_file_throws
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
void __vybeMain() async {
  if (!FileSystemEntity.isWatchSupported) return;
  final file = File('does_not_exist_watch.txt');
  try {
    final stream = file.watch();
    stream.listen((_) {}, onError: (e) {
      __p('Error on stream');
    });
  } catch (e) {
    __p('FileSystemException thrown');
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('');
}
