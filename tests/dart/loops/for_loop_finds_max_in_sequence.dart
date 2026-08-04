// vybe-test: dart/loops/for_loop_finds_max_in_sequence
// origin: languages/dart/tests/dart/test_loops.rs

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
  var nums = [3, 9, 1, 7, 4];
  var max = nums[0];
  for (var i = 1; i < nums.length; i++) {
    if (nums[i] > max) max = nums[i];
  }
  __p(max);
}

void main() {
  __vybeMain();
  __check('9');
}
