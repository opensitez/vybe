// vybe-test: dart/unary_operator_overload/unary_bitwise_not_twice_restores
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

class Toggle {
  int t;
  Toggle(this.t);
  Toggle operator ~() {
    return Toggle(~t);
  }
}
void __vybeMain() {
  __p((~(~Toggle(99))).t);
}

void main() {
  __vybeMain();
  __check('99');
}
