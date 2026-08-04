// vybe-test: dart/closures/closure_for_each_accumulates_with_capture
// origin: languages/dart/tests/dart/test_closures.rs

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
  var total = 0;
  [1, 2, 3, 4].forEach((n) {
    total += n;
  });
  __p(total);
}

void main() {
  __vybeMain();
  __check('10');
}
