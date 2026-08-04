// vybe-test: dart/operator_overloading/operator_plus_returns_new_instance
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

class Pair {
  int a;
  int b;
  Pair(this.a, this.b);
  Pair operator +(Pair o) {
    return Pair(a + o.a, b + o.b);
  }
}
void __vybeMain() {
  var p = Pair(1, 1) + Pair(2, 3);
  __p(p.a);
}

void main() {
  __vybeMain();
  __check('3');
}
