// vybe-test: dart/iterable_methods/iterable_fold_after_where
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
  Iterable<int> nums = [1, 2, 3, 4, 5, 6];
  __p(nums.where((n) => n % 2 == 0).fold(0, (acc, n) => acc + n));
}

void main() {
  __vybeMain();
  __check('12');
}
