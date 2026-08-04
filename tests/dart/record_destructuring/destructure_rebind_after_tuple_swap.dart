// vybe-test: dart/record_destructuring/destructure_rebind_after_tuple_swap
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
  var (m, n) = (1, 2);
  (m, n) = (n, m);
  var (p, q) = (m, n);
  __p(p);
  __p(q);
}

void main() {
  __vybeMain();
  __check('2\n1');
}
