// vybe-test: dart/getters_setters/arrow_body_getter_returns_expression
// origin: languages/dart/tests/dart/test_getters_setters.rs

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
  int x = 2;
  int y = 3;
  int get sum => x + y;
}
void __vybeMain() {
  __p(Point().sum);
}

void main() {
  __vybeMain();
  __check('5');
}
