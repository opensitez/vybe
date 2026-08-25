// vybe-test: dart/dart_isolate_transferable_data/send_port_send_send_port
// origin: languages/dart/tests/dart/test_dart_isolate_transferable_data.rs

import 'dart:async';
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
  final innerPort = ReceivePort();
  port.send(innerPort.sendPort);
  innerPort.listen((msg) {
    if (msg == 'ping') port.send('pong');
  });
}
// `await` of a `void` expression is a compile error under dart 3.10.4, so the
// async scaffold must answer a Future for `main` to await (measured).
Future<void> __vybeMain() async {
  final receivePort = ReceivePort();
  await Isolate.spawn(isolateMain, receivePort.sendPort);
  
  // Damaged test repaired: `first` consumes the port's single subscription,
  // so the original `take(2)` afterwards threw "Bad state: Stream has
  // already been listened to" under dart 3.10.4 (measured). ONE listen keeps
  // the round-trip intent: the inner SendPort arrives, 'ping' goes through
  // it, and the 'pong' reply lands on the same receivePort.
  final completer = Completer();
  receivePort.listen((msg) {
    if (msg is SendPort) {
      msg.send('ping');
    } else {
      completer.complete(msg);
    }
  });
  final pong = await completer.future;
  __p(pong);
}

Future<void> main() async {
  await __vybeMain();
  __check('pong');
}
