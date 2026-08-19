// vybe-test: dart/dart_isolate_capabilities_pause_kill/isolate_add_on_exit_listener_response
// origin: languages/dart/tests/dart/test_dart_isolate_capabilities_pause_kill.rs

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

void isolateMain(_) {}
void __vybeMain() async {
  final isolate = await Isolate.spawn(isolateMain, null);
  final port = ReceivePort();
  isolate.addOnExitListener(port.sendPort, response: 'custom_exit');
  final msg = await port.first;
  __p(msg);
}

Future<void> main() async {
  await __vybeMain();
  __check('custom_exit');
}
