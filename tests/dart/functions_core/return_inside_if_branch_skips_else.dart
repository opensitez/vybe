// vybe-test: dart/functions_core/return_inside_if_branch_skips_else
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

String classify(int n) {
  if (n < 0) {
    return 'negative';
  }
  if (n == 0) {
    return 'zero';
  }
  return 'positive';
}
void __vybeMain() {
  __p(classify(-1));
  __p(classify(0));
  __p(classify(3));
}

void main() {
  __vybeMain();
  __check('negative\nzero\npositive');
}
