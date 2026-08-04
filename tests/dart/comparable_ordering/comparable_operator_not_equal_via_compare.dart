// vybe-test: dart/comparable_ordering/comparable_operator_not_equal_via_compare
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

class Id implements Comparable<Id> {
  int id;
  Id(this.id);
  int compareTo(Id other) => id.compareTo(other.id);
}
void __vybeMain() {
  __p(Id(1).compareTo(Id(2)) != 0);
}

void main() {
  __vybeMain();
  __check('true');
}
