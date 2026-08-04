// vybe-test: dart/unary_operator_overload/unary_minus_returns_new_instance
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

class Pair {
  int a;
  Pair(this.a);
  Pair operator -() {
    return Pair(-a);
  }
}
void __vybeMain() {
  var p = Pair(2);
  var q = -p;
  __p(p.a);
  __p(q.a);
}

void main() {
  __vybeMain();
  __check('2\n-2');
}
