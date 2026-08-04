// vybe-test: dart/generics_core/typed_list_int_literal
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
  List<int> nums = [1, 2, 3];
  __p(nums.length);
  __p(nums[1]);
}

void main() {
  __vybeMain();
  __check('3\n2');
}
