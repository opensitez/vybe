// vybe-test: dart/typedefs_core/typedef_int_binary_function_sum
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

typedef IntOp = int Function(int, int);
int add(int a, int b) {
  return a + b;
}
void __vybeMain() {
  IntOp op = add;
  __p(op(7, 5));
}

void main() {
  __vybeMain();
  __check('12');
}
