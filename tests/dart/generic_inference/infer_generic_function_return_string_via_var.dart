// vybe-test: dart/generic_inference/infer_generic_function_return_string_via_var
// origin: languages/dart/tests/dart/test_generic_inference.rs

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

List<T> singleton<T>(T value) {
  return [value];
}
void __vybeMain() {
  var list = singleton('only');
  __p(list.first);
}

void main() {
  __vybeMain();
  __check('only');
}
