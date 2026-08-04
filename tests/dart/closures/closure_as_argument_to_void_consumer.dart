// vybe-test: dart/closures/closure_as_argument_to_void_consumer
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

void consume(int x, void Function(int) fn) {
  fn(x);
}
void __vybeMain() {
  var seen = 0;
  consume(7, (n) {
    seen = n;
  });
  __p(seen);
}

void main() {
  __vybeMain();
  __check('7');
}
