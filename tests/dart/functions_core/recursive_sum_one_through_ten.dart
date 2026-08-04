// vybe-test: dart/functions_core/recursive_sum_one_through_ten
// origin: languages/dart/tests/dart/test_functions_core.rs

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

int sum(int n) {
  if (n <= 0) {
    return 0;
  }
  return n + sum(n - 1);
}
void __vybeMain() {
  __p(sum(10));
}

void main() {
  __vybeMain();
  __check('55');
}
