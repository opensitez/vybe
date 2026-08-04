// vybe-test: dart/unary_operator_overload/unary_bitwise_not_in_conditional
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

class Gate {
  int g;
  Gate(this.g);
  Gate operator ~() {
    return Gate(~g);
  }
}
void __vybeMain() {
  var x = ~Gate(0);
  __p(x.g == -1);
}

void main() {
  __vybeMain();
  __check('true');
}
