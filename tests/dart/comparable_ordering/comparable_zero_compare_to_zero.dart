// vybe-test: dart/comparable_ordering/comparable_zero_compare_to_zero
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

class Zero implements Comparable<Zero> {
  int v;
  Zero(this.v);
  int compareTo(Zero other) => v.compareTo(other.v);
}
void __vybeMain() {
  __p(Zero(0).compareTo(Zero(0)));
}

void main() {
  __vybeMain();
  __check('0');
}
