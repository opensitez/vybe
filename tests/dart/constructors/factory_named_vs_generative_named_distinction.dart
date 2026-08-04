// vybe-test: dart/constructors/factory_named_vs_generative_named_distinction
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

class Point {
  int x;
  int y;
  Point(this.x, this.y);
  Point.fromXY(int a, int b) : x = a, y = b;
  factory Point.middle() {
    return Point(50, 50);
  }
}
void __vybeMain() {
  __p(Point.middle().x);
}

void main() {
  __vybeMain();
  __check('50');
}
