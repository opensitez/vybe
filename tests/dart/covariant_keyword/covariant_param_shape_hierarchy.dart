// vybe-test: dart/covariant_keyword/covariant_param_shape_hierarchy
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

class Shape {}
class Circle extends Shape {
  int r;
  Circle(this.r);
}
class Drawer {
  void draw(Shape s) {}
}
class CircleDrawer extends Drawer {
  @override
  void draw(covariant Circle c) {
    __p(c.r);
  }
}
void __vybeMain() {
  CircleDrawer().draw(Circle(5));
}

void main() {
  __vybeMain();
  __check('5');
}
