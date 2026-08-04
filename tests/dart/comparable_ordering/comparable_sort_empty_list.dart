// vybe-test: dart/comparable_ordering/comparable_sort_empty_list
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

class E implements Comparable<E> {
  int e;
  E(this.e);
  int compareTo(E other) => e.compareTo(other.e);
}
void __vybeMain() {
  var list = <E>[];
  list.sort();
  __p(list.length);
}

void main() {
  __vybeMain();
  __check('0');
}
