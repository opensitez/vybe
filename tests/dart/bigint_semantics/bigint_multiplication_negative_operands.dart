// vybe-test: dart/bigint_semantics/bigint_multiplication_negative_operands
// origin: languages/dart/tests/dart/test_bigint_semantics.rs

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
  var a = BigInt.from(-6);
  var b = BigInt.from(7);
  __p(a * b);
}

void main() {
  __vybeMain();
  __check('-42');
}
