// vybe-test: dart/dart_isolate_capabilities_pause_kill/isolate_ping_with_custom_response
// origin: languages/dart/tests/dart/test_dart_isolate_capabilities_pause_kill.rs

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
  // Keep it alive
  Timer(Duration(hours: 1), () {});
}
void __vybeMain() async {
  final receivePort = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.ping(receivePort.sendPort, response: 42);
  
  final msg = await receivePort.first;
  __p(msg);
  isolate.kill();
}

Future<void> main() async {
  await __vybeMain();
  __check('42');
}
