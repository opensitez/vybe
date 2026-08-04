// vybe-test: dart/break_continue/for_break_on_last_valid_index
// origin: languages/dart/tests/dart/test_break_continue.rs

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
  var nums = [10, 20, 30];
  var picked = 0;
  for (var i = 0; i < nums.length; i++) {
    picked = nums[i];
    if (i == nums.length - 1) break;
  }
  __p(picked);
}

void main() {
  __vybeMain();
  __check('30');
}
