// vybe-test: dart/enum_enhanced/enhanced_enum_two_int_fields_sum
// origin: languages/dart/tests/dart/test_enum_enhanced.rs

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

enum Point {
  origin(0, 0),
  unit(1, 1);
  final int x;
  final int y;
  const Point(this.x, this.y);
}
void __vybeMain() {
  __p(Point.unit.x + Point.unit.y);
}

void main() {
  __vybeMain();
  __check('2');
}
