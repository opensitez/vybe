// vybe-test: dart/extensions/int_static_sum_three_integers
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

extension IntAgg on int {
  static int sum3(int a, int b, int c) => a + b + c;
}
void __vybeMain() {
  __p(IntAgg.sum3(4, 5, 6));
}

void main() {
  __vybeMain();
  __check('15');
}
