// vybe-test: dart/comparable_ordering/comparable_compare_negative_one_less
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

class N implements Comparable<N> {
  int n;
  N(this.n);
  int compareTo(N other) => n.compareTo(other.n);
}
void __vybeMain() {
  __p(N(-10).compareTo(N(-5)) < 0);
}

void main() {
  __vybeMain();
  __check('true');
}
