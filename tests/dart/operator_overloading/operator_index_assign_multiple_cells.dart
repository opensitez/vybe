// vybe-test: dart/operator_overloading/operator_index_assign_multiple_cells
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

class Matrix {
  List<int> flat;
  Matrix(this.flat);
  int operator [](int i) => flat[i];
  void operator []=(int i, int v) {
    flat[i] = v;
  }
}
void __vybeMain() {
  var m = Matrix([0, 0, 0]);
  m[0] = 1;
  m[2] = 3;
  __p(m[0] + m[2]);
}

void main() {
  __vybeMain();
  __check('4');
}
