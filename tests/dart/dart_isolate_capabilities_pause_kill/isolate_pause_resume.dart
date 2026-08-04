// vybe-test: dart/dart_isolate_capabilities_pause_kill/isolate_pause_resume
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
void __vybeMain() async {
  // We can't pause current isolate and resume it easily without a deadlock,
  // but we can spawn one, pause it, and resume it.
  final receivePort = ReceivePort();
  final isolate = await Isolate.spawn((port) {
    (port as SendPort).send('started');
  }, receivePort.sendPort);
  
  final cap = isolate.pause();
  isolate.resume(cap);
  
  final msg = await receivePort.first;
  print(msg);
}

Future<void> main() async {
  await __vybeMain();
  __check('started');
}
