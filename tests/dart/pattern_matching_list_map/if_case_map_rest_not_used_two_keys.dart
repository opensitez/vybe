// vybe-test: dart/pattern_matching_list_map/if_case_map_rest_not_used_two_keys
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
  var m = {'p': 1, 'q': 2, 'r': 3};
  if (m case {'p': var x, 'r': var z}) {
    __p(x);
    __p(z);
  } else {
    __p(0);
  }
}

void main() {
  __vybeMain();
  __check('1\n3');
}
