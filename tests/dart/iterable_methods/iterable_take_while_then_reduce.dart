// vybe-test: dart/iterable_methods/iterable_take_while_then_reduce
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
  Iterable<int> nums = [2, 4, 6, 1, 3];
  __p(nums.takeWhile((n) => n % 2 == 0).reduce((a, b) => a + b));
}

void main() {
  __vybeMain();
  __check('12');
}
