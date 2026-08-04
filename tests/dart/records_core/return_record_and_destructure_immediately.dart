// vybe-test: dart/records_core/return_record_and_destructure_immediately
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

(int, int) origin() {
  return (0, 0);
}
void __vybeMain() {
  var (x, y) = origin();
  __p(x);
  __p(y);
}

void main() {
  __vybeMain();
  __check('0\n0');
}
