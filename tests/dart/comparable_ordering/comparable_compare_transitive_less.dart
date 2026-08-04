// vybe-test: dart/comparable_ordering/comparable_compare_transitive_less
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

class T implements Comparable<T> {
  int t;
  T(this.t);
  int compareTo(T other) => t.compareTo(other.t);
}
void __vybeMain() {
  var a = T(1);
  var b = T(2);
  var c = T(3);
  __p(a.compareTo(b) < 0 && b.compareTo(c) < 0);
}

void main() {
  __vybeMain();
  __check('true');
}
