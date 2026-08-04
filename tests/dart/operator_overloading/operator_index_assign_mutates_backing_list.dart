// vybe-test: dart/operator_overloading/operator_index_assign_mutates_backing_list
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

class Grid {
  List<int> data;
  Grid(this.data);
  int operator [](int i) {
    return data[i];
  }
  void operator []=(int i, int v) {
    data[i] = v;
  }
}
void __vybeMain() {
  var g = Grid([1, 2, 3]);
  g[1] = 99;
  __p(g.data[1]);
}

void main() {
  __vybeMain();
  __check('99');
}
