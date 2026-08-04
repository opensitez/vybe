// vybe-test: dart/const_deep/static_const_used_in_const_expression
// origin: languages/dart/tests/dart/test_const_deep.rs

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

class Limits {
  static const int max = 100;
}
void __vybeMain() {
  const half = Limits.max ~/ 2;
  __p(half);
}

void main() {
  __vybeMain();
  __check('50');
}
