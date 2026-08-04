// vybe-test: dart/comparable_ordering/comparable_negative_values_order
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

class Temp implements Comparable<Temp> {
  int celsius;
  Temp(this.celsius);
  int compareTo(Temp other) => celsius.compareTo(other.celsius);
}
void __vybeMain() {
  __p(Temp(-5).compareTo(Temp(0)));
}

void main() {
  __vybeMain();
  __check('-1');
}
