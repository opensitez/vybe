// vybe-test: dart/class_modifiers/base_class_method_called_on_subclass_reference
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

base class Shape {
  int sides() {
    return 0;
  }
}
class Square extends Shape {
  @override
  int sides() {
    return 4;
  }
}
void __vybeMain() {
  Shape s = Square();
  __p(s.sides());
}

void main() {
  __vybeMain();
  __check('4');
}
