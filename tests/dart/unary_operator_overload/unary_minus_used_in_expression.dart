// vybe-test: dart/unary_operator_overload/unary_minus_used_in_expression
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

class Point {
  int x;
  Point(this.x);
  Point operator -() {
    return Point(-x);
  }
}
void __vybeMain() {
  __p((-Point(4)).x + 1);
}

void main() {
  __vybeMain();
  __check('-3');
}
