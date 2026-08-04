// vybe-test: dart/extensions/int_static_absolute_difference
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

extension IntDiff on int {
  static int absDiff(int a, int b) => a > b ? a - b : b - a;
}
void __vybeMain() {
  __p(IntDiff.absDiff(5, 12));
}

void main() {
  __vybeMain();
  __check('7');
}
