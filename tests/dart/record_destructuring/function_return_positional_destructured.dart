// vybe-test: dart/record_destructuring/function_return_positional_destructured
// origin: languages/dart/tests/dart/test_record_destructuring.rs

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

(int, int) pair() => (7, 8);
void __vybeMain() {
  var (x, y) = pair();
  __p(x);
  __p(y);
}

void main() {
  __vybeMain();
  __check('7\n8');
}
