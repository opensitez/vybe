// vybe-test: dart/dart_isolate_errors_exit/isolate_errors_are_fatal_behavior
// origin: languages/dart/tests/dart/test_dart_isolate_errors_exit.rs

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

void isolateMain(SendPort port) {
  Isolate.current.setErrorsFatal(true);
  try {
    throw Exception('test_error');
  } catch (e) {
    port.send('caught');
  }
}
void __vybeMain() async {
  final port = ReceivePort();
  await Isolate.spawn(isolateMain, port.sendPort);
  final msg = await port.first;
  __p(msg);
}

Future<void> main() async {
  await __vybeMain();
  __check('caught');
}
