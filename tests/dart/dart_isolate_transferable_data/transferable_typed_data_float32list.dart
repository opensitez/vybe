// vybe-test: dart/dart_isolate_transferable_data/transferable_typed_data_float32list
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

void __vybeMain() {
  final l = Float32List(2);
  l[0] = 3.5;
  final ttd = TransferableTypedData.fromList([l.buffer.asUint8List()]);
  // Damaged test repaired: `materialize()` answers a ByteBuffer (no
  // `getFloat32` — did not compile under dart 3.10.4); the legal read goes
  // through `.asByteData()`, which prints 3.5 (measured).
  final bd = ttd.materialize().asByteData();
  __p(bd.getFloat32(0, Endian.host));
}

void main() {
  __vybeMain();
  __check('3.5');
}
