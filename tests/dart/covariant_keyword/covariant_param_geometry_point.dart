// vybe-test: dart/covariant_keyword/covariant_param_geometry_point
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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
}
class ColoredPoint extends Point {
  String color;
  ColoredPoint(int x, int y, this.color) : super(x, y);
}
class Plotter {
  void mark(Point p) {}
}
class ColorPlotter extends Plotter {
  @override
  void mark(covariant ColoredPoint p) {
    __p('${p.x},${p.color}');
  }
}
void __vybeMain() {
  ColorPlotter().mark(ColoredPoint(1, 2, 'red'));
}

void main() {
  __vybeMain();
  __check('1,red');
}
