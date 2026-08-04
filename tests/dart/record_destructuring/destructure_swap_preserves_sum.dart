// vybe-test: dart/record_destructuring/destructure_swap_preserves_sum
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

void __vybeMain() {
  var x = 4;
  var y = 6;
  var before = x + y;
  (x, y) = (y, x);
  __p(x + y == before);
}

void main() {
  __vybeMain();
  __check('true');
}
