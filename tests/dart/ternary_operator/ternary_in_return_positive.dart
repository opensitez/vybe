// vybe-test: dart/ternary_operator/ternary_in_return_positive
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

int absVal(int n) {
  return n >= 0 ? n : -n;
}
void __vybeMain() {
  __p(absVal(5));
}

void main() {
  __vybeMain();
  __check('5');
}
