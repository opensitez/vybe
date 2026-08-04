// vybe-test: dart/function_apply/apply_named_negative_int
// origin: languages/dart/tests/dart/test_function_apply.rs

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

int adjust(int n, {int delta = 0}) {
  return n + delta;
}
void __vybeMain() {
  __p(Function.apply(adjust, [10], {#delta: -3}));
}

void main() {
  __vybeMain();
  __check('7');
}
