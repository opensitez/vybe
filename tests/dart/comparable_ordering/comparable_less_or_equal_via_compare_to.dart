// vybe-test: dart/comparable_ordering/comparable_less_or_equal_via_compare_to
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

class Size implements Comparable<Size> {
  int n;
  Size(this.n);
  int compareTo(Size other) => n.compareTo(other.n);
  bool operator <=(Size other) => compareTo(other) <= 0;
}
void __vybeMain() {
  __p(Size(4) <= Size(4));
}

void main() {
  __vybeMain();
  __check('true');
}
