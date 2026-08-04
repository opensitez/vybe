// vybe-test: dart/ternary_operator/ternary_chained_three_levels
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
  var n = 15;
  var size = n > 20 ? 'xl' : n > 10 ? 'lg' : n > 5 ? 'md' : 'sm';
  __p(size);
}

void main() {
  __vybeMain();
  __check('lg');
}
