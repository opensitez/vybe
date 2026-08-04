// vybe-test: dart/ternary_operator/ternary_with_negative_number
// origin: languages/dart/tests/dart/test_ternary_operator.rs

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
  var n = -4;
  __p(n < 0 ? 'neg' : 'pos');
}

void main() {
  __vybeMain();
  __check('neg');
}
