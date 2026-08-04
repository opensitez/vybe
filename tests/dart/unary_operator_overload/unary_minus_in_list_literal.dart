// vybe-test: dart/unary_operator_overload/unary_minus_in_list_literal
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

class N {
  int v;
  N(this.v);
  N operator -() {
    return N(-v);
  }
}
void __vybeMain() {
  var items = [-N(1), -N(2)];
  __p(items[0].v + items[1].v);
}

void main() {
  __vybeMain();
  __check('-3');
}
