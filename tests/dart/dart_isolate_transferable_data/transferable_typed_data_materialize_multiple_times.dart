// vybe-test: dart/dart_isolate_transferable_data/transferable_typed_data_materialize_multiple_times
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
  final list = Uint8List.fromList([5, 6, 7]);
  final ttd = TransferableTypedData.fromList([list]);
  final m1 = ttd.materialize();
  final m2 = ttd.materialize();
  // Multiple materialize calls usually return the exact same ByteData instance or a new one pointing to same memory.
  // But actually the memory is moved back to the current isolate context.
  __p(m1.lengthInBytes == m2.lengthInBytes);
}

void main() {
  __vybeMain();
  __check('true');
}
