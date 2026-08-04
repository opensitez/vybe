// vybe-test: dart/comparable_ordering/comparable_compare_equal_returns_zero
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

class Rank implements Comparable<Rank> {
  int level;
  Rank(this.level);
  int compareTo(Rank other) => level.compareTo(other.level);
}
void __vybeMain() {
  __p(Rank(5).compareTo(Rank(5)));
}

void main() {
  __vybeMain();
  __check('0');
}
