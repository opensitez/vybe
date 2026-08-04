// vybe-test: dart/loops/for_loop_reverses_into_new_list
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
  var src = [1, 2, 3];
  var rev = <int>[];
  for (var i = src.length - 1; i >= 0; i--) {
    rev.add(src[i]);
  }
  __p(rev.join('-'));
}

void main() {
  __vybeMain();
  __check('3-2-1');
}
