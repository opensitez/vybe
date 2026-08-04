// vybe-test: dart/loops/while_loop_sentinel_minus_one
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
  var data = [5, 3, 8, -1, 2];
  var idx = 0;
  var sum = 0;
  while (data[idx] != -1) {
    sum += data[idx];
    idx++;
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('16');
}
