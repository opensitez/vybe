// vybe-test: dart/generic_inference/infer_apply_with_string_callback
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

T apply<T>(T val, T Function(T) fn) {
  return fn(val);
}
void __vybeMain() {
  __p(apply('hi', (s) => s + s));
}

void main() {
  __vybeMain();
  __check('hihi');
}
