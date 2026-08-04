// vybe-test: dart/comparable_ordering/comparable_sort_two_elements_swap
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

class Duo implements Comparable<Duo> {
  int d;
  Duo(this.d);
  int compareTo(Duo other) => d.compareTo(other.d);
}
void __vybeMain() {
  var list = [Duo(2), Duo(1)];
  list.sort();
  __p(list[0].d);
  __p(list[1].d);
}

void main() {
  __vybeMain();
  __check('1\n2');
}
