// vybe-test: dart/pattern_matching_list_map/switch_map_when_two_field_sum
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
  var m = {'u': 2, 'v': 3};
  __p(switch (m) {
    {'u': var u, 'v': var v} when u + v == 5 => 'sum5',
    _ => 'other' });
}

void main() {
  __vybeMain();
  __check('sum5');
}
