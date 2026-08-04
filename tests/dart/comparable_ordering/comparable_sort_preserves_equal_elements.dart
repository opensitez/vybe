// vybe-test: dart/comparable_ordering/comparable_sort_preserves_equal_elements
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

class Key implements Comparable<Key> {
  int k;
  Key(this.k);
  int compareTo(Key other) => k.compareTo(other.k);
}
void __vybeMain() {
  var list = [Key(2), Key(1), Key(2)];
  list.sort();
  __p(list[0].k);
  __p(list[1].k);
  __p(list[2].k);
}

void main() {
  __vybeMain();
  __check('1\n2\n2');
}
