// vybe-test: dart/operator_overloading/operator_plus_double_application
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

class Inc {
  int n;
  Inc(this.n);
  Inc operator +(Inc o) {
    return Inc(n + o.n);
  }
}
void __vybeMain() {
  var base = Inc(1);
  var step = Inc(4);
  __p((base + step + step).n);
}

void main() {
  __vybeMain();
  __check('9');
}
