// vybe-test: dart/loops/for_loop_counts_matching_elements
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
  var nums = [1, 2, 2, 3, 2, 4];
  var count = 0;
  for (var i = 0; i < nums.length; i++) {
    if (nums[i] == 2) count++;
  }
  __p(count);
}

void main() {
  __vybeMain();
  __check('3');
}
