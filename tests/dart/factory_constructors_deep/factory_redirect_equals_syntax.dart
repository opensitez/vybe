// vybe-test: dart/factory_constructors_deep/factory_redirect_equals_syntax
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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

class Point {
  int x;
  int y;
  Point(this.x, this.y);
  Point._zero() : x = 0, y = 0;
  factory Point.zero() = Point._zero;
}
void __vybeMain() {
  __p(Point.zero().x);
}

void main() {
  __vybeMain();
  __check('0');
}
