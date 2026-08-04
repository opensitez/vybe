// vybe-test: dart/record_destructuring/sync_return_destructure_from_bounds_helper
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

({int min, int max}) bounds(List<int> xs) {
  return (min: xs.first, max: xs.last);
}
void __vybeMain() {
  var (min: lo, max: hi) = bounds([3, 9, 1, 7]);
  __p(lo);
  __p(hi);
}

void main() {
  __vybeMain();
  __check('3\n7');
}
