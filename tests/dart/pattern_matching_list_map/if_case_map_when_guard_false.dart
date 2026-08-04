// vybe-test: dart/pattern_matching_list_map/if_case_map_when_guard_false
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
  var m = {'n': 0};
  if (m case {'n': var v} when v > 0) {
    __p('yes');
  } else {
    __p('no');
  }
}

void main() {
  __vybeMain();
  __check('no');
}
