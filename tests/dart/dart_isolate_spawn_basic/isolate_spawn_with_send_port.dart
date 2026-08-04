// vybe-test: dart/dart_isolate_spawn_basic/isolate_spawn_with_send_port
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
void isolateMain(SendPort port) {
  port.send('message_from_isolate');
}
void __vybeMain() async {
  final receivePort = ReceivePort();
  await Isolate.spawn(isolateMain, receivePort.sendPort);
  final message = await receivePort.first;
  __p(message);
}

Future<void> main() async {
  await __vybeMain();
  __check('message_from_isolate');
}
