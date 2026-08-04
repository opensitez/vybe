// vybe-test: dart/factory_constructors_deep/factory_selects_subclass_by_type_code
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

class Shape {
  int sides;
  Shape(this.sides);
  factory Shape.fromCode(String code) {
    if (code == 'tri') {
      return Triangle();
    }
    return Square();
  }
}
class Triangle extends Shape {
  Triangle() : super(3);
}
class Square extends Shape {
  Square() : super(4);
}
void __vybeMain() {
  __p(Shape.fromCode('tri').sides);
}

void main() {
  __vybeMain();
  __check('3');
}
