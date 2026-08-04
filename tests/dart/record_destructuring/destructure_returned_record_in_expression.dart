// vybe-test: dart/record_destructuring/destructure_returned_record_in_expression
// origin: languages/dart/tests/dart/test_record_destructuring.rs

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

(int, int) twice(int n) => (n, n * 2);
void __vybeMain() {
  var (a, b) = twice(5);
  __p(a + b);
}

void main() {
  __vybeMain();
  __check('15');
}
