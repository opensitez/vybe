// vybe-test: dart/bigint_semantics/bigint_power_small
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
  var base = BigInt.from(2);
  __p(base * base * base);
}

void main() {
  __vybeMain();
  __check('8');
}
