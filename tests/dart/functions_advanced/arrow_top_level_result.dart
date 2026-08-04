// vybe-test: dart/functions_advanced/arrow_top_level_result
// origin: languages/dart/tests/dart/test_functions_advanced.rs

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

int sq(int x) => x * x; void __vybeMain() { __p(sq(9)); }

void main() {
  __vybeMain();
  __check('81');
}
