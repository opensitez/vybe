// vybe-test: dart/extensions/int_static_is_within_inclusive_range
// origin: languages/dart/tests/dart/test_extensions.rs

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

extension IntRange on int {
  static bool inRange(int n, int lo, int hi) => n >= lo && n <= hi;
}
void __vybeMain() {
  __p(IntRange.inRange(7, 1, 10));
}

void main() {
  __vybeMain();
  __check('true');
}
