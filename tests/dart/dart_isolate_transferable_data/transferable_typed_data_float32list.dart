// vybe-test: dart/dart_isolate_transferable_data/transferable_typed_data_float32list
// origin: languages/dart/tests/dart/test_dart_isolate_transferable_data.rs

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
import 'dart:typed_data';
void __vybeMain() {
  final l = Float32List(2);
  l[0] = 3.5;
  final ttd = TransferableTypedData.fromList([l.buffer.asUint8List()]);
  final bd = ttd.materialize();
  __p(bd.getFloat32(0, Endian.host));
}

void main() {
  __vybeMain();
  __check('3.5');
}
