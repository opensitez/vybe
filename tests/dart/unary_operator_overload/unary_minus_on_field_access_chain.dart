// vybe-test: dart/unary_operator_overload/unary_minus_on_field_access_chain
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

class Outer {
  Inner inner;
  Outer(this.inner);
}
class Inner {
  int v;
  Inner(this.v);
  Inner operator -() {
    return Inner(-v);
  }
}
void __vybeMain() {
  __p((-Outer(Inner(5)).inner).v);
}

void main() {
  __vybeMain();
  __check('-5');
}
