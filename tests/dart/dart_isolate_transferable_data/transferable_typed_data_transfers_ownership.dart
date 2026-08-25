// vybe-test: dart/dart_isolate_transferable_data/transferable_typed_data_transfers_ownership
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
  // Damaged test repaired, twice over: the prints went to `print` instead of
  // `__p`, so `__check` compared against an empty buffer no matter what
  // happened; and under dart 3.10.4 `fromList` COPIES — `list[0]` stays
  // readable and this prints "accessible" (measured), never "detached".
  try {
    list[0];
    __p('accessible');
  } catch(e) {
    __p('detached');
  }
}

void main() {
  __vybeMain();
  __check('accessible');
}
