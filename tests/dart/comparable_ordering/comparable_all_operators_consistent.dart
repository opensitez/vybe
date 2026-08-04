// vybe-test: dart/comparable_ordering/comparable_all_operators_consistent
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

class Ord implements Comparable<Ord> {
  int o;
  Ord(this.o);
  int compareTo(Ord other) => o.compareTo(other.o);
  bool operator <(Ord x) => compareTo(x) < 0;
  bool operator <=(Ord x) => compareTo(x) <= 0;
  bool operator >(Ord x) => compareTo(x) > 0;
  bool operator >=(Ord x) => compareTo(x) >= 0;
}
void __vybeMain() {
  var a = Ord(2);
  var b = Ord(5);
  __p(a < b);
  __p(a <= b);
  __p(b > a);
  __p(b >= a);
}

void main() {
  __vybeMain();
  __check('true\ntrue\ntrue\ntrue');
}
