// vybe-test: dart/null_safety_advanced/assert_non_null_int
// origin: languages/dart/tests/dart/test_null_safety_advanced.rs

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

void __vybeMain() { int? n = 42; __p(n! + 1); }

void main() {
  __vybeMain();
  __check('43');
}
