// vybe-test: dart/generic_inference/infer_cast_list_after_generic_map
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

List<R> convert<T, R>(List<T> items, R Function(T) fn) {
  return items.map(fn).toList();
}
void __vybeMain() {
  var lengths = convert(['ab', 'cde'], (s) => s.length);
  __p(lengths.join(','));
}

void main() {
  __vybeMain();
  __check('2,3');
}
