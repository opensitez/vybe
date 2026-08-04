// vybe-test: dart/loops/nested_for_loops_count_cartesian_pairs
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
  var count = 0;
  for (var r = 0; r < 2; r++) {
    for (var c = 0; c < 3; c++) {
      count++;
    }
  }
  __p(count);
}

void main() {
  __vybeMain();
  __check('6');
}
