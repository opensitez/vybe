// vybe-test: dart/dart_isolate_ports_messaging/raw_receive_port_handler
// origin: languages/dart/tests/dart/test_dart_isolate_ports_messaging.rs

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
void __vybeMain() async {
  final completer = Completer();
  final port = RawReceivePort((msg) {
    completer.complete(msg);
  });
  port.sendPort.send('raw_message');
  final msg = await completer.future;
  __p(msg);
  port.close();
}

Future<void> main() async {
  await __vybeMain();
  __check('raw_message');
}
