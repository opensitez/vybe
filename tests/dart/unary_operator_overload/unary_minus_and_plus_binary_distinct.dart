// vybe-test: dart/unary_operator_overload/unary_minus_and_plus_binary_distinct
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

class AddNeg {
  int n;
  AddNeg(this.n);
  AddNeg operator -() {
    return AddNeg(-n);
  }
  AddNeg operator +(AddNeg o) {
    return AddNeg(n + o.n);
  }
}
void __vybeMain() {
  var a = AddNeg(5);
  __p((-a).n);
  __p((a + AddNeg(1)).n);
}

void main() {
  __vybeMain();
  __check('-5\n6');
}
