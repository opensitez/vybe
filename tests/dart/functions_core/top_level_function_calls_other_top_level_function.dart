// vybe-test: dart/functions_core/top_level_function_calls_other_top_level_function
// origin: languages/dart/tests/dart/test_functions_core.rs

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

int doubleIt(int x) {
  return x * 2;
}
int quadrupleIt(int x) {
  return doubleIt(doubleIt(x));
}
void __vybeMain() {
  __p(quadrupleIt(4));
}

void main() {
  __vybeMain();
  __check('16');
}
