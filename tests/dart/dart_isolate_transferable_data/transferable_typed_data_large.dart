// vybe-test: dart/dart_isolate_transferable_data/transferable_typed_data_large
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
  final l1 = Uint8List(100000);
  final ttd = TransferableTypedData.fromList([l1]);
  final bd = ttd.materialize();
  __p(bd.lengthInBytes);
}

void main() {
  __vybeMain();
  __check('100000');
}
