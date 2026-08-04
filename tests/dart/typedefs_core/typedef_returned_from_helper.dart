// vybe-test: dart/typedefs_core/typedef_returned_from_helper
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

typedef Op = int Function(int);
Op makeAdder(int delta) {
  return (int n) => n + delta;
}
void __vybeMain() {
  Op add5 = makeAdder(5);
  __p(add5(10));
}

void main() {
  __vybeMain();
  __check('15');
}
