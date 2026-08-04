// vybe-test: dart/function_apply/apply_positional_bool_or
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

bool eitherTrue(bool a, bool b) {
  return a || b;
}
void __vybeMain() {
  __p(Function.apply(eitherTrue, [true, false]));
}

void main() {
  __vybeMain();
  __check('true');
}
