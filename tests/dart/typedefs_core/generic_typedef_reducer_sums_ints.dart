// vybe-test: dart/typedefs_core/generic_typedef_reducer_sums_ints
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

typedef Reducer<T> = T Function(T, T);
int foldLeft(List<int> items, Reducer<int> combine, int seed) {
  var acc = seed;
  for (var item in items) {
    acc = combine(acc, item);
  }
  return acc;
}
void __vybeMain() {
  __p(foldLeft([1, 2, 3], (a, b) => a + b, 0));
}

void main() {
  __vybeMain();
  __check('6');
}
