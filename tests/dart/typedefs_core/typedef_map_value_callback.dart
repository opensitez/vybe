// vybe-test: dart/typedefs_core/typedef_map_value_callback
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

typedef Callback = int Function(int);
int runAll(Map<String, Callback> table, int seed) {
  var total = 0;
  table.forEach((key, fn) {
    total += fn(seed);
  });
  return total;
}
void __vybeMain() {
  var table = <String, Callback>{
    'a': (n) => n + 1,
    'b': (n) => n * 2,
  };
  __p(runAll(table, 3));
}

void main() {
  __vybeMain();
  __check('10');
}
