// vybe-test: dart/dart_isolate_errors_exit/isolate_kill_immediate_does_not_trigger_exit_listener
// origin: languages/dart/tests/dart/test_dart_isolate_errors_exit.rs

import 'dart:isolate';
import 'dart:async';

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

void isolateMain(_) {
  while(true) {}
}
void __vybeMain() async {
  final exitPort = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.addOnExitListener(exitPort.sendPort, response: 'exit_detected');
  
  // Immediate kill might or might not trigger exit listener depending on VM internals.
  // Actually, killing an isolate SHOULD trigger the exit listener.
  // We'll test if it does.
  isolate.kill(priority: Isolate.immediate);
  final msg = await exitPort.first;
  print(msg);
}

Future<void> main() async {
  await __vybeMain();
  __check('exit_detected');
}
