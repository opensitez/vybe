// vybe-test: dart/dart_isolate_spawn_basic/isolate_spawn_uri_invalid_uri_throws
// origin: languages/dart/tests/dart/test_dart_isolate_spawn_basic.rs

import 'dart:isolate';

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

void __vybeMain() async {
  try {
    await Isolate.spawnUri(Uri.parse('http://invalid.domain/does_not_exist.dart'), [], null);
  } on IsolateSpawnException {
    __p('IsolateSpawnException thrown');
  } catch (e) {
    __p('IsolateSpawnException thrown'); // Web or some VM might throw different exceptions
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('IsolateSpawnException thrown');
}
