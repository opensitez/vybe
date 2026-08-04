// vybe-test: dart/loops/nested_for_loops_print_row_column_indices
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
  for (var r = 0; r < 2; r++) {
    for (var c = 0; c < 2; c++) {
      __p('$r$c');
    }
  }
}

void main() {
  __vybeMain();
  __check('00\n01\n10\n11');
}
