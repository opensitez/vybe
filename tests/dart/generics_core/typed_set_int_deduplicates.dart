// vybe-test: dart/generics_core/typed_set_int_deduplicates
// origin: languages/dart/tests/dart/test_generics_core.rs

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
  Set<int> nums = {1, 2, 2, 3};
  __p(nums.length);
  __p(nums.contains(2));
}

void main() {
  __vybeMain();
  __check('3\ntrue');
}
