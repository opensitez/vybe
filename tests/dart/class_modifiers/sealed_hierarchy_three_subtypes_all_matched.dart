// vybe-test: dart/class_modifiers/sealed_hierarchy_three_subtypes_all_matched
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

sealed class Shape {}
class Circle extends Shape {
  int r;
  Circle(this.r);
}
class Square extends Shape {
  int side;
  Square(this.side);
}
class Triangle extends Shape {
  int base;
  Triangle(this.base);
}
int measure(Shape s) {
  switch (s) {
    case Circle(r: var radius):
      return radius;
    case Square(side: var s):
      return s;
    case Triangle(base: var b):
      return b;
  }
}
void __vybeMain() {
  __p(measure(Square(6)));
}

void main() {
  __vybeMain();
  __check('6');
}
