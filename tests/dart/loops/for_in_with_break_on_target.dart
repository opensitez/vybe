// vybe-test: dart/loops/for_in_with_break_on_target
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
  var found = 0;
  for (var x in [2, 4, 6, 8, 10]) {
    if (x == 6) {
      found = x;
      break;
    }
  }
  __p(found);
}

void main() {
  __vybeMain();
  __check('6');
}
