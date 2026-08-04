// vybe-test: dart/dart_isolate_spawn_basic/isolate_exit
// origin: languages/dart/tests/dart/test_dart_isolate_spawn_basic.rs

import 'dart:isolate';
void isolateMain(SendPort port) {
  Isolate.exit(port, 'exited');
  // shouldn't execute
  port.send('unreachable');
}
void main() async {
  final receivePort = ReceivePort();
  await Isolate.spawn(isolateMain, receivePort.sendPort);
  final message = await receivePort.first;
  print(message);
}
