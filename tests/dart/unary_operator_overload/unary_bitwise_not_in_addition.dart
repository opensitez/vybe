// vybe-test: dart/unary_operator_overload/unary_bitwise_not_in_addition
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

class Flag {
  int f;
  Flag(this.f);
  Flag operator ~() {
    return Flag(~f);
  }
}
void __vybeMain() {
  __p((~Flag(0)).f + 1);
}

void main() {
  __vybeMain();
  __check('0');
}
