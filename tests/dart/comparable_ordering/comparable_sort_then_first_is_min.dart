// vybe-test: dart/comparable_ordering/comparable_sort_then_first_is_min
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

class Min implements Comparable<Min> {
  int v;
  Min(this.v);
  int compareTo(Min other) => v.compareTo(other.v);
}
void __vybeMain() {
  var list = [Min(5), Min(-1), Min(3)];
  list.sort();
  __p(list.first.v);
}

void main() {
  __vybeMain();
  __check('-1');
}
