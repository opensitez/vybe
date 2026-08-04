// vybe-test: dart/comparison_operators/type_promoted_int_comparison_after_is_check
// origin: languages/dart/tests/dart/test_comparison_operators.rs

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
  Object? value = 8;
  if (value is int) {
    __p(value > 5);
  }
}

void main() {
  __vybeMain();
  __check('true');
}
