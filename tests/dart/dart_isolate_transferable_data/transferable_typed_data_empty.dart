// vybe-test: dart/dart_isolate_transferable_data/transferable_typed_data_empty
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
  final ttd = TransferableTypedData.fromList([]);
  final bd = ttd.materialize();
  __p(bd.lengthInBytes);
}

void main() {
  __vybeMain();
  __check('0');
}
