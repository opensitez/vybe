// vybe-test: dart/unary_operator_overload/unary_minus_and_bitwise_not_combined
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

class Dual {
  int v;
  Dual(this.v);
  Dual operator -() {
    return Dual(-v);
  }
  Dual operator ~() {
    return Dual(~v);
  }
}
void __vybeMain() {
  __p((-Dual(3)).v);
  __p((~Dual(3)).v);
}

void main() {
  __vybeMain();
  __check('-3\n-4');
}
