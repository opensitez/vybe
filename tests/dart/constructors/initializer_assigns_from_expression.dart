// vybe-test: dart/constructors/initializer_assigns_from_expression
// origin: languages/dart/tests/dart/test_constructors.rs

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

class Square {
  int side;
  int area;
  Square(int s) : side = s, area = s * s;
}
void __vybeMain() {
  __p(Square(6).area);
}

void main() {
  __vybeMain();
  __check('36');
}
