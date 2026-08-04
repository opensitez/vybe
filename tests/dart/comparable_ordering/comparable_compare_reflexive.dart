// vybe-test: dart/comparable_ordering/comparable_compare_reflexive
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

class Ref implements Comparable<Ref> {
  int r;
  Ref(this.r);
  int compareTo(Ref other) => r.compareTo(other.r);
}
void __vybeMain() {
  var x = Ref(7);
  __p(x.compareTo(x));
}

void main() {
  __vybeMain();
  __check('0');
}
