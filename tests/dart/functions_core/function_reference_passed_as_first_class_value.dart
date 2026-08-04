// vybe-test: dart/functions_core/function_reference_passed_as_first_class_value
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

int doubleIt(int x) => x * 2;
int applyTwice(int x, int Function(int) fn) {
  return fn(fn(x));
}
void __vybeMain() {
  __p(applyTwice(3, doubleIt));
}

void main() {
  __vybeMain();
  __check('12');
}
