// vybe-test: dart/field_initializers/initializer_list_computes_from_constructor_params
// origin: languages/dart/tests/dart/test_field_initializers.rs

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

class Circle {
  int radius;
  int diameter;
  Circle(int r) : radius = r, diameter = r * 2;
}
void __vybeMain() {
  __p(Circle(6).diameter);
}

void main() {
  __vybeMain();
  __check('12');
}
