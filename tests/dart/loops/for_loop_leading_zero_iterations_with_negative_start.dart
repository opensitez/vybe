// vybe-test: dart/loops/for_loop_leading_zero_iterations_with_negative_start
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
  var ran = 0;
  for (var i = -3; i < -3; i++) {
    ran++;
  }
  __p(ran);
}

void main() {
  __vybeMain();
  __check('0');
}
