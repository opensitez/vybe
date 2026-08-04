// vybe-test: dart/typedefs_core/typedef_int_unary_function_invoked
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

typedef IntFn = int Function(int);
int triple(int n) {
  return n * 3;
}
void __vybeMain() {
  IntFn fn = triple;
  __p(fn(4));
}

void main() {
  __vybeMain();
  __check('12');
}
