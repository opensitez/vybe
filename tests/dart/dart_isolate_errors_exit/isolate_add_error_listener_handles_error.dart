// vybe-test: dart/dart_isolate_errors_exit/isolate_add_error_listener_handles_error
// origin: languages/dart/tests/dart/test_dart_isolate_errors_exit.rs

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
import 'dart:async';
void isolateMain(_) {
  throw Exception('isolate_error');
}
void __vybeMain() async {
  final port = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.addErrorListener(port.sendPort);
  
  final msg = await port.first;
  __p((msg[0] as String).contains('isolate_error'));
  port.close();
}

Future<void> main() async {
  await __vybeMain();
  __check('true');
}
