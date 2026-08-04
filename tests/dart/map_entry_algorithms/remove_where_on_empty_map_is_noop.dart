// vybe-test: dart/map_entry_algorithms/remove_where_on_empty_map_is_noop
// origin: languages/dart/tests/dart/test_map_entry_algorithms.rs

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
  var m = <String, int>{};
  m.removeWhere((k, v) => true);
  __p(m.isEmpty);
  __p(m.length);
}

void main() {
  __vybeMain();
  __check('true\n0');
}
