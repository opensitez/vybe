// vybe-test: dart/map_core/map_from_iterables_zips_parallel_key_value_lists
// origin: languages/dart/tests/dart/test_map_core.rs

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
  var m = Map.fromIterables(['a', 'b', 'c'], [10, 20, 30]);
  __p(m['a']);
  __p(m['c']);
  __p(m.length);
}

void main() {
  __vybeMain();
  __check('10\n30\n3');
}
