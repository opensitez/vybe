// vybe-test: dart/records_core/record_passed_through_two_functions
// origin: languages/dart/tests/dart/test_records_core.rs

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

(int, int) makePair() {
  return (3, 5);
}
int doubleFirst((int, int) p) {
  return p.$1 * 2;
}
void __vybeMain() {
  __p(doubleFirst(makePair()));
}

void main() {
  __vybeMain();
  __check('6');
}
