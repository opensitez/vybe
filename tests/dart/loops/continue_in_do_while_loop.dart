// vybe-test: dart/loops/continue_in_do_while_loop
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
  var i = 0;
  do {
    i++;
    if (i == 2) continue;
    __p(i);
  } while (i < 4);
}

void main() {
  __vybeMain();
  __check('1\n3\n4');
}
