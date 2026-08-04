// vybe-test: dart/operator_overloading/operator_index_read_first_slot
// origin: languages/dart/tests/dart/test_operator_overloading.rs

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

class Row {
  List<int> cells;
  Row(this.cells);
  int operator [](int i) {
    return cells[i];
  }
}
void __vybeMain() {
  __p(Row([10, 20, 30])[0]);
}

void main() {
  __vybeMain();
  __check('10');
}
