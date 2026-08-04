// vybe-test: dart/comparable_ordering/comparable_less_equal_mixed_operators
// origin: languages/dart/tests/dart/test_comparable_ordering.rs

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

class Point implements Comparable<Point> {
  int x;
  Point(this.x);
  int compareTo(Point other) => x.compareTo(other.x);
  bool operator <(Point o) => compareTo(o) < 0;
  bool operator <=(Point o) => compareTo(o) <= 0;
}
void __vybeMain() {
  __p(Point(1) < Point(2));
  __p(Point(2) <= Point(2));
}

void main() {
  __vybeMain();
  __check('true\ntrue');
}
