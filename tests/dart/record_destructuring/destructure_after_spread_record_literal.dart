// vybe-test: dart/record_destructuring/destructure_after_spread_record_literal
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

void __vybeMain() {
  var base = (a: 1, b: 2);
  var (a: x, b: y, c: z) = (a: base.a, b: base.b, c: 3);
  __p(x);
  __p(y);
  __p(z);
}

void main() {
  __vybeMain();
  __check('1\n2\n3');
}
