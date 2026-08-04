// vybe-test: dart/field_initializers/named_constructor_initializer_list_sets_fields
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

class Vector {
  int x;
  int y;
  Vector(this.x, this.y);
  Vector.zero() : x = 0, y = 0;
}
void __vybeMain() {
  __p(Vector.zero().y);
}

void main() {
  __vybeMain();
  __check('0');
}
