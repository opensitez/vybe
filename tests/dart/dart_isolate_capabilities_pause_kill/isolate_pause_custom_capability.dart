// vybe-test: dart/dart_isolate_capabilities_pause_kill/isolate_pause_custom_capability
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
void isolateMain(SendPort port) {
  port.send('started');
}
void __vybeMain() async {
  final receivePort = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, receivePort.sendPort);
  
  final cap = Capability();
  isolate.pause(cap);
  isolate.resume(cap);
  
  final msg = await receivePort.first;
  __p(msg);
}

Future<void> main() async {
  await __vybeMain();
  __check('started');
}
