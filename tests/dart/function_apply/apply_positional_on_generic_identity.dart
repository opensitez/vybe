// vybe-test: dart/function_apply/apply_positional_on_generic_identity
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

T id<T>(T v) {
  return v;
}
void __vybeMain() {
  __p(Function.apply(id, [99]));
}

void main() {
  __vybeMain();
  __check('99');
}
