// vybe-test: dart/dart_isolate_capabilities_pause_kill/isolate_ping
// origin: languages/dart/tests/dart/test_dart_isolate_capabilities_pause_kill.rs

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

void isolateMain(_) {}
void __vybeMain() async {
  final receivePort = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.ping(receivePort.sendPort, response: 'pong');
  
  final msg = await receivePort.first;
  __p(msg);
  // Actually, ping response might arrive before or after isolate exits.
  // Because it's an empty isolate, it exits fast, so ping might return the 'pong' or exit event depending.
  // It should be 'pong' though.
}

Future<void> main() async {
  await __vybeMain();
  __check('pong');
}
