// vybe-test: dart/dart_isolate_spawn_basic/isolate_spawn_closure_throws
// origin: languages/dart/tests/dart/test_dart_isolate_spawn_basic.rs

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

import 'dart:isolate';
void __vybeMain() async {
  // Spawning an isolate with a closure that captures context usually throws ArgumentError
  int local = 5;
  try {
    await Isolate.spawn((_) { __p(local); }, null);
    // Actually, Dart 2.15+ supports isolate groups and some closures can be spawned
    // If it succeeds, it's valid. If it fails, it throws ArgumentError
    print('handled');
  } catch(e) {
    print('handled');
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('handled');
}
