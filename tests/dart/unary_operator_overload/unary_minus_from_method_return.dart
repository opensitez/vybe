// vybe-test: dart/unary_operator_overload/unary_minus_from_method_return
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

class Source {
  int n;
  Source(this.n);
  Source operator -() {
    return Source(-n);
  }
  Source make() {
    return Source(4);
  }
}
void __vybeMain() {
  __p((-Source(0).make()).n);
}

void main() {
  __vybeMain();
  __check('-4');
}
