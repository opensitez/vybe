// vybe-test: dart/dart_isolate_ports_messaging/send_port_send_complex_object
// origin: languages/dart/tests/dart/test_dart_isolate_ports_messaging.rs

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

void __vybeMain() async {
  final port = ReceivePort();
  final map = {'a': 1, 'b': [2, 3]};
  port.sendPort.send(map);
  final msg = await port.first;
  __p(msg['b'][1]);
}

Future<void> main() async {
  await __vybeMain();
  __check('3');
}
