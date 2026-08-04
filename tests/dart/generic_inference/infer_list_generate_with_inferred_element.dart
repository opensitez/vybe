// vybe-test: dart/generic_inference/infer_list_generate_with_inferred_element
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

List<T> duplicate<T>(T value, int count) {
  return List.generate(count, (_) => value);
}
void __vybeMain() {
  var items = duplicate(true, 2);
  __p(items[0]);
  __p(items[1]);
}

void main() {
  __vybeMain();
  __check('true\ntrue');
}
