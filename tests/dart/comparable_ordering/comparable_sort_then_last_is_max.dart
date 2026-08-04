// vybe-test: dart/comparable_ordering/comparable_sort_then_last_is_max
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

class Max implements Comparable<Max> {
  int v;
  Max(this.v);
  int compareTo(Max other) => v.compareTo(other.v);
}
void __vybeMain() {
  var list = [Max(5), Max(-1), Max(3)];
  list.sort();
  __p(list.last.v);
}

void main() {
  __vybeMain();
  __check('5');
}
