// vybe-test: dart/dart_isolate_transferable_data/send_port_send_send_port
// origin: languages/dart/tests/dart/test_dart_isolate_transferable_data.rs

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
  final innerPort = ReceivePort();
  port.send(innerPort.sendPort);
  innerPort.listen((msg) {
    if (msg == 'ping') port.send('pong');
  });
}
void __vybeMain() async {
  final receivePort = ReceivePort();
  await Isolate.spawn(isolateMain, receivePort.sendPort);
  
  final innerSendPort = await receivePort.first as SendPort;
  innerSendPort.send('ping');
  
  // listen for pong on the same receivePort (or a new one if isolate sent it there)
  // Wait, the isolate sent the first msg, and then sends 'pong' to 'port'
  // So receivePort will get another message
  final list = await receivePort.take(2).toList();
  __p(list[1]);
}

Future<void> main() async {
  await __vybeMain();
  __check('pong');
}
