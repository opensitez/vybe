// vybe-test: dart/comparable_ordering/comparable_median_of_three
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

class Mid implements Comparable<Mid> {
  int m;
  Mid(this.m);
  int compareTo(Mid other) => m.compareTo(other.m);
}
Mid median(Mid a, Mid b, Mid c) {
  var list = [a, b, c];
  list.sort();
  return list[1];
}
void __vybeMain() {
  __p(median(Mid(3), Mid(1), Mid(2)).m);
}

void main() {
  __vybeMain();
  __check('2');
}
