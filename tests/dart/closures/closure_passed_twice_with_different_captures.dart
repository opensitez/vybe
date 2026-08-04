// vybe-test: dart/closures/closure_passed_twice_with_different_captures
// origin: languages/dart/tests/dart/test_closures.rs

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

int run(int Function(int) fn, int x) {
  return fn(x);
}
void __vybeMain() {
  var offset = 3;
  __p(run((n) => n + offset, 10));
  offset = 5;
  __p(run((n) => n + offset, 10));
}

void main() {
  __vybeMain();
  __check('13\n15');
}
