// vybe-test: dart/dart_isolate_transferable_data/send_port_send_transfer_capability
// origin: languages/dart/tests/dart/test_dart_isolate_transferable_data.rs

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
  final cap = Capability();
  port.send(cap);
}
void __vybeMain() async {
  final receivePort = ReceivePort();
  await Isolate.spawn(isolateMain, receivePort.sendPort);
  final msg = await receivePort.first;
  __p(msg is Capability);
}

Future<void> main() async {
  await __vybeMain();
  __check('true');
}
