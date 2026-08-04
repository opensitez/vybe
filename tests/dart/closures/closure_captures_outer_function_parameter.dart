// vybe-test: dart/closures/closure_captures_outer_function_parameter
// origin: languages/dart/tests/dart/test_closures.rs

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

int wrap(int base) {
  var add = (int n) => base + n;
  return add(4);
}
void __vybeMain() {
  __p(wrap(10));
}

void main() {
  __vybeMain();
  __check('14');
}
