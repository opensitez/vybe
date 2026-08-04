// vybe-test: dart/loops/nested_while_counts_grid_cells
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
  var rows = 2;
  var cols = 4;
  var r = 0;
  var total = 0;
  while (r < rows) {
    var c = 0;
    while (c < cols) {
      total++;
      c++;
    }
    r++;
  }
  __p(total);
}

void main() {
  __vybeMain();
  __check('8');
}
