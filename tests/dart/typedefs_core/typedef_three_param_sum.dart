// vybe-test: dart/typedefs_core/typedef_three_param_sum
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

typedef Sum3 = int Function(int, int, int);
int sum3(int a, int b, int c) {
  return a + b + c;
}
void __vybeMain() {
  Sum3 fn = sum3;
  __p(fn(1, 2, 3));
}

void main() {
  __vybeMain();
  __check('6');
}
