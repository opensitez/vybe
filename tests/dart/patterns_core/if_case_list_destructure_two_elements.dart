// vybe-test: dart/patterns_core/if_case_list_destructure_two_elements
// origin: languages/dart/tests/dart/test_patterns_core.rs

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
  var xs = [10, 20];
  if (xs case [var a, var b]) {
    __p(a);
    __p(b);
  }
}

void main() {
  __vybeMain();
  __check('10\n20');
}
