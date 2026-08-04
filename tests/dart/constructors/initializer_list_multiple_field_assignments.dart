// vybe-test: dart/constructors/initializer_list_multiple_field_assignments
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

class Line {
  int x1;
  int y1;
  int x2;
  int y2;
  Line(int a, int b, int c, int d)
      : x1 = a, y1 = b, x2 = c, y2 = d;
}
void __vybeMain() {
  var l = Line(0, 0, 3, 4);
  __p(l.x2 + l.y2);
}

void main() {
  __vybeMain();
  __check('7');
}
