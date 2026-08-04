// vybe-test: dart/abstract_members/abstract_method_implemented_by_subclass
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Shape {
  double area();
}
class Square extends Shape {
  int side;
  Square(this.side);
  double area() {
    return side * side;
  }
}
void __vybeMain() {
  __p(Square(4).area());
}

void main() {
  __vybeMain();
  __check('16.0');
}
