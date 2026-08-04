// vybe-test: dart/break_continue/for_in_continue_skips_zeros_in_list
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
  var sum = 0;
  for (var x in [1, 0, 2, 0, 3]) {
    if (x == 0) continue;
    sum += x;
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('6');
}
