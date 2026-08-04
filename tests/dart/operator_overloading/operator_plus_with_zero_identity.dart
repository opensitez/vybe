// vybe-test: dart/operator_overloading/operator_plus_with_zero_identity
// origin: languages/dart/tests/dart/test_operator_overloading.rs

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

class Num {
  int v;
  Num(this.v);
  Num operator +(Num other) {
    return Num(v + other.v);
  }
}
void __vybeMain() {
  var n = Num(5) + Num(0);
  __p(n.v);
}

void main() {
  __vybeMain();
  __check('5');
}
