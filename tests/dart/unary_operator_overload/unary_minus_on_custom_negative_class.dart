// vybe-test: dart/unary_operator_overload/unary_minus_on_custom_negative_class
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

class Negative {
  int value;
  Negative(this.value);
  Negative operator -() {
    return Negative(-value);
  }
  bool isNegative() {
    return value < 0;
  }
}
void __vybeMain() {
  var n = Negative(-9);
  var pos = -n;
  __p(pos.value);
  __p(pos.isNegative());
}

void main() {
  __vybeMain();
  __check('9\nfalse');
}
