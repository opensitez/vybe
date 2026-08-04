// vybe-test: dart/comparable_ordering/comparable_find_min_via_compare_to
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

class Val implements Comparable<Val> {
  int n;
  Val(this.n);
  int compareTo(Val other) => n.compareTo(other.n);
}
Val minOf(Val a, Val b) {
  return a.compareTo(b) <= 0 ? a : b;
}
void __vybeMain() {
  __p(minOf(Val(3), Val(7)).n);
}

void main() {
  __vybeMain();
  __check('3');
}
