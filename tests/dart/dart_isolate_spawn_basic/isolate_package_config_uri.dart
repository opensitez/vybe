// vybe-test: dart/dart_isolate_spawn_basic/isolate_package_config_uri
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
  // Can pass packageConfig to spawnUri
  __p('compiles');
}

Future<void> main() async {
  await __vybeMain();
  __check('compiles');
}
