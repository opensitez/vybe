// vybe-test: dart/typedefs_core/generic_typedef_box_list_nested_length
// origin: languages/dart/tests/dart/test_typedefs_core.rs

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

typedef BoxList<T> = List<List<T>>;
void __vybeMain() {
  BoxList<int> matrix = [
    [1, 2],
    [3],
  ];
  __p(matrix.length);
  __p(matrix[0].length);
  __p(matrix[1].single);
}

void main() {
  __vybeMain();
  __check('2\n2\n3');
}
