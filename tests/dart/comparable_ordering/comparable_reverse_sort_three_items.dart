// vybe-test: dart/comparable_ordering/comparable_reverse_sort_three_items
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

class Num implements Comparable<Num> {
  int n;
  Num(this.n);
  int compareTo(Num other) => n.compareTo(other.n);
}
void __vybeMain() {
  var list = [Num(1), Num(2), Num(3)];
  list.sort((a, b) => b.compareTo(a));
  __p(list.map((e) => e.n).join(','));
}

void main() {
  __vybeMain();
  __check('3,2,1');
}
