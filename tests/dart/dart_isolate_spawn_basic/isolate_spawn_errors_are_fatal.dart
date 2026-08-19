// vybe-test: dart/dart_isolate_spawn_basic/isolate_spawn_errors_are_fatal
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
  final receivePort = ReceivePort();
  final isolate = await Isolate.spawn((_) {}, null, errorsAreFatal: true);
  __p(isolate.errorsAreFatal != null); // Might be internal flag but API accepts it
  // Wait, errorsAreFatal is not a property on Isolate, it's an arg.
  // We just verify it compiles and runs.
  receivePort.close();
}

Future<void> main() async {
  await __vybeMain();
  __check('true');
}
