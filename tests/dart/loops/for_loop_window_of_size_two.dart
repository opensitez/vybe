// vybe-test: dart/loops/for_loop_window_of_size_two
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
  var nums = [1, 2, 3, 4];
  var windows = 0;
  for (var i = 0; i < nums.length - 1; i++) {
    windows++;
  }
  __p(windows);
}

void main() {
  __vybeMain();
  __check('3');
}
