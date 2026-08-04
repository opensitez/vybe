// vybe-test: dart/iterable_methods/iterable_to_set_deduplicates_elements
// origin: languages/dart/tests/dart/test_iterable_methods.rs

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
  Iterable<int> nums = [1, 2, 2, 3, 3, 3];
  var s = nums.toSet();
  __p(s.length);
  __p(s.contains(2));
}

void main() {
  __vybeMain();
  __check('3\ntrue');
}
