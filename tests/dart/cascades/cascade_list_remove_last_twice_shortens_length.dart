// vybe-test: dart/cascades/cascade_list_remove_last_twice_shortens_length
// origin: languages/dart/tests/dart/test_cascades.rs

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
  var nums = [1, 2, 3, 4];
  nums..removeLast()..removeLast();
  __p(nums.length);
  __p(nums.join(','));
}

void main() {
  __vybeMain();
  __check('2\n1,2');
}
