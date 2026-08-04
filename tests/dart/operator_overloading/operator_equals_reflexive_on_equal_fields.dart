// vybe-test: dart/operator_overloading/operator_equals_reflexive_on_equal_fields
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

class RGB {
  int r;
  int g;
  int b;
  RGB(this.r, this.g, this.b);
  bool operator ==(Object other) {
    if (other is! RGB) return false;
    return r == other.r && g == other.g && b == other.b;
  }
}
void __vybeMain() {
  var c = RGB(1, 2, 3);
  __p(c == RGB(1, 2, 3));
}

void main() {
  __vybeMain();
  __check('true');
}
