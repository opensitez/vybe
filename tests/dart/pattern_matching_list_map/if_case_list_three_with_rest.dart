// vybe-test: dart/pattern_matching_list_map/if_case_list_three_with_rest
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
  var xs = [1, 2, 3, 4];
  if (xs case [var a, ...var rest]) {
    __p(a);
    __p(rest.join(','));
  } else {
    __p('no');
  }
}

void main() {
  __vybeMain();
  __check('1\n2,3,4');
}
