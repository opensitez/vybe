// vybe-test: dart/pattern_matching_list_map/if_case_list_pair_destructure
// origin: languages/dart/tests/dart/test_pattern_matching_list_map.rs

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
    __p(a + b);
  } else {
    __p(0);
  }
}

void main() {
  __vybeMain();
  __check('30');
}
