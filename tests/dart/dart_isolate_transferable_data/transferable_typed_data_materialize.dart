// vybe-test: dart/dart_isolate_transferable_data/transferable_typed_data_materialize
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
  final list = Uint8List.fromList([10, 20, 30]);
  final ttd = TransferableTypedData.fromList([list]);
  // Damaged test repaired: `materialize()` answers a ByteBuffer, which has no
  // `getUint8` — the original spelling did not compile under dart 3.10.4.
  // The legal read goes through `.asByteData()`, and the single interpolated
  // print produces "3:20" (measured), not the "3\n20" the check wanted.
  final materialized = ttd.materialize().asByteData();
  __p('${materialized.lengthInBytes}:${materialized.getUint8(1)}');
}

void main() {
  __vybeMain();
  __check('3:20');
}
