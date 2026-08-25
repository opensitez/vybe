// vybe-test: dart/dart_isolate_transferable_data/isolate_send_transferable_typed_data
// origin: languages/dart/tests/dart/test_dart_isolate_transferable_data.rs

import 'dart:isolate';
import 'dart:typed_data';

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
  final list = Uint8List.fromList([99, 100]);
  final ttd = TransferableTypedData.fromList([list]);
  port.send(ttd);
}
// `await` of a `void` expression is a compile error under dart 3.10.4, so the
// async scaffold must answer a Future for `main` to await (measured).
Future<void> __vybeMain() async {
  final receivePort = ReceivePort();
  await Isolate.spawn(isolateMain, receivePort.sendPort);
  
  final msg = await receivePort.first;
  // Damaged test repaired: `materialize()` answers a ByteBuffer (no
  // `getUint8` — did not compile under dart 3.10.4); the legal read goes
  // through `.asByteData()`.
  final bd = (msg as TransferableTypedData).materialize().asByteData();
  __p(bd.getUint8(0));
}

Future<void> main() async {
  await __vybeMain();
  __check('99');
}
