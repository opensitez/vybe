// vybe-test: dart/comparable_ordering/comparable_operator_greater_false_when_equal
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

class Box implements Comparable<Box> {
  int v;
  Box(this.v);
  int compareTo(Box other) => v.compareTo(other.v);
  bool operator >(Box other) => compareTo(other) > 0;
}
void __vybeMain() {
  __p(Box(5) > Box(5));
}

void main() {
  __vybeMain();
  __check('false');
}
