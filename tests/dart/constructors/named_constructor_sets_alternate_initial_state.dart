// vybe-test: dart/constructors/named_constructor_sets_alternate_initial_state
// origin: languages/dart/tests/dart/test_constructors.rs

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
  Point.origin() : x = 0, y = 0;
}
void __vybeMain() {
  var p = Point.origin();
  __p(p.x);
}

void main() {
  __vybeMain();
  __check('0');
}
