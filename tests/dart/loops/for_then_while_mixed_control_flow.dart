// vybe-test: dart/loops/for_then_while_mixed_control_flow
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
  var sum = 0;
  for (var i = 1; i <= 3; i++) {
    sum += i;
  }
  var j = 0;
  while (j < 2) {
    sum += 10;
    j++;
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('26');
}
