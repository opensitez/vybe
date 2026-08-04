// vybe-test: dart/const_deep/const_list_of_const_objects
// origin: languages/dart/tests/dart/test_const_deep.rs

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

class Cell {
  final int v;
  const Cell(this.v);
}
void __vybeMain() {
  const row = [Cell(1), Cell(2), Cell(3)];
  __p(row[1].v);
}

void main() {
  __vybeMain();
  __check('2');
}
