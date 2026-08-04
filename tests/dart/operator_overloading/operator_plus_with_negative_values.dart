// vybe-test: dart/operator_overloading/operator_plus_with_negative_values
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

class Signed {
  int v;
  Signed(this.v);
  Signed operator +(Signed o) {
    return Signed(v + o.v);
  }
}
void __vybeMain() {
  __p((Signed(-2) + Signed(5)).v);
}

void main() {
  __vybeMain();
  __check('3');
}
