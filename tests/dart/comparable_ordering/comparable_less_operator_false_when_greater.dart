// vybe-test: dart/comparable_ordering/comparable_less_operator_false_when_greater
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

class W implements Comparable<W> {
  int w;
  W(this.w);
  int compareTo(W other) => w.compareTo(other.w);
  bool operator <(W other) => compareTo(other) < 0;
}
void __vybeMain() {
  __p(W(9) < W(1));
}

void main() {
  __vybeMain();
  __check('false');
}
