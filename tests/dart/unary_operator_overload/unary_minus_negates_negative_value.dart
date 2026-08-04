// vybe-test: dart/unary_operator_overload/unary_minus_negates_negative_value
// origin: languages/dart/tests/dart/test_unary_operator_overload.rs

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

class Signed {
  int v;
  Signed(this.v);
  Signed operator -() {
    return Signed(-v);
  }
}
void __vybeMain() {
  __p((-Signed(-3)).v);
}

void main() {
  __vybeMain();
  __check('3');
}
