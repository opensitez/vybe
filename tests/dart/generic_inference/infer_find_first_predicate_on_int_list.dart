// vybe-test: dart/generic_inference/infer_find_first_predicate_on_int_list
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

T findFirst<T>(List<T> items, bool Function(T) test) {
  return items.firstWhere(test);
}
void __vybeMain() {
  __p(findFirst([1, 2, 3, 4], (n) => n > 2));
}

void main() {
  __vybeMain();
  __check('3');
}
