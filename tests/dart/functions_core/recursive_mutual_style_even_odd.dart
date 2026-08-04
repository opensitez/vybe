// vybe-test: dart/functions_core/recursive_mutual_style_even_odd
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

bool isEven(int n) {
  if (n == 0) {
    return true;
  }
  return isOdd(n - 1);
}
bool isOdd(int n) {
  if (n == 0) {
    return false;
  }
  return isEven(n - 1);
}
void __vybeMain() {
  __p(isEven(4));
  __p(isOdd(4));
}

void main() {
  __vybeMain();
  __check('true\nfalse');
}
