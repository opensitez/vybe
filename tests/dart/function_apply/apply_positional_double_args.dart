// vybe-test: dart/function_apply/apply_positional_double_args
// origin: languages/dart/tests/dart/test_function_apply.rs

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

double avg(double a, double b) {
  return (a + b) / 2;
}
void __vybeMain() {
  __p(Function.apply(avg, [2.0, 4.0]));
}

void main() {
  __vybeMain();
  __check('3.0');
}
