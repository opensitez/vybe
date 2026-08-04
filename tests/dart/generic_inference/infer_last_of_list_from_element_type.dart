// vybe-test: dart/generic_inference/infer_last_of_list_from_element_type
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

T lastOf<T>(List<T> items) {
  return items.last;
}
void __vybeMain() {
  __p(lastOf([1, 2, 3]));
}

void main() {
  __vybeMain();
  __check('3');
}
