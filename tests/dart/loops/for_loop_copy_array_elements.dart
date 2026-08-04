// vybe-test: dart/loops/for_loop_copy_array_elements
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
  var src = [4, 5, 6];
  var dst = List<int>.filled(src.length, 0);
  for (var i = 0; i < src.length; i++) {
    dst[i] = src[i];
  }
  __p(dst[1]);
}

void main() {
  __vybeMain();
  __check('5');
}
