// vybe-test: dart/generic_inference/infer_comparable_min_from_string_args
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

T minOf<T extends Comparable<T>>(T a, T b) {
  return a.compareTo(b) <= 0 ? a : b;
}
void __vybeMain() {
  __p(minOf('zebra', 'ant'));
}

void main() {
  __vybeMain();
  __check('ant');
}
