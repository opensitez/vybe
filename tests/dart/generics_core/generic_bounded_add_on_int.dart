// vybe-test: dart/generics_core/generic_bounded_add_on_int
// origin: languages/dart/tests/dart/test_generics_core.rs

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

T addNums<T extends num>(T a, T b) {
  return (a + b) as T;
}
void __vybeMain() {
  __p(addNums(3, 4));
}

void main() {
  __vybeMain();
  __check('7');
}
