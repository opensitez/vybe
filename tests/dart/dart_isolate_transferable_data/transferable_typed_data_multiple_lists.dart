// vybe-test: dart/dart_isolate_transferable_data/transferable_typed_data_multiple_lists
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
  final l1 = Uint8List.fromList([1, 2]);
  final l2 = Uint8List.fromList([3, 4]);
  // TransferableTypedData concatenates the byte data of the lists
  final ttd = TransferableTypedData.fromList([l1, l2]);
  final bd = ttd.materialize();
  __p(bd.lengthInBytes);
  __p(bd.getUint8(2)); // Should be 3
}

void main() {
  __vybeMain();
  __check('4\n3');
}
