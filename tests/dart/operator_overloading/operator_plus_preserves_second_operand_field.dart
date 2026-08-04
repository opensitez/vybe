// vybe-test: dart/operator_overloading/operator_plus_preserves_second_operand_field
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

class Mix {
  int x;
  int y;
  Mix(this.x, this.y);
  Mix operator +(Mix o) {
    return Mix(x + o.x, y);
  }
}
void __vybeMain() {
  __p((Mix(1, 9) + Mix(2, 0)).y);
}

void main() {
  __vybeMain();
  __check('9');
}
