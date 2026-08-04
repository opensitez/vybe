// vybe-test: dart/dart_typed_data_float_lists/float32list_precision_loss
// origin: languages/dart/tests/dart/test_dart_typed_data_float_lists.rs

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

import 'dart:typed_data';
void __vybeMain() {
  // Float32 loses precision on double
  final f = Float32List(1);
  f[0] = 3.141592653589793238; // double precision
  __p(f[0] == 3.141592653589793238); // false
  __p(f[0] != 0.0);
}

void main() {
  __vybeMain();
  __check('false\ntrue');
}
