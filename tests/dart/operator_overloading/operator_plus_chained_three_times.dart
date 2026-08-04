// vybe-test: dart/operator_overloading/operator_plus_chained_three_times
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

class Adder {
  int n;
  Adder(this.n);
  Adder operator +(Adder other) {
    return Adder(n + other.n);
  }
}
void __vybeMain() {
  var r = Adder(1) + Adder(2) + Adder(3);
  __p(r.n);
}

void main() {
  __vybeMain();
  __check('6');
}
