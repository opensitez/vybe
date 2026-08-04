// vybe-test: dart/field_initializers/initializer_list_sets_both_coords_from_single_param
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

class Point1D {
  int x;
  int y;
  Point1D(int v) : x = v, y = v;
}
void __vybeMain() {
  __p(Point1D(8).x + Point1D(8).y);
}

void main() {
  __vybeMain();
  __check('16');
}
