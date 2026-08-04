// vybe-test: dart/comparable_ordering/comparable_sort_single_element
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

class Solo implements Comparable<Solo> {
  int x;
  Solo(this.x);
  int compareTo(Solo other) => x.compareTo(other.x);
}
void __vybeMain() {
  var list = [Solo(42)];
  list.sort();
  __p(list[0].x);
}

void main() {
  __vybeMain();
  __check('42');
}
