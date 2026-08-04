// vybe-test: dart/record_destructuring/double_swap_restores_original_values
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
  var p = 5;
  var q = 15;
  (p, q) = (q, p);
  (p, q) = (q, p);
  __p(p);
  __p(q);
}

void main() {
  __vybeMain();
  __check('5\n15');
}
