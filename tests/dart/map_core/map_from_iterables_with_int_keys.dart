// vybe-test: dart/map_core/map_from_iterables_with_int_keys
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
  var m = Map<int, String>.fromIterables([0, 1, 2], ['zero', 'one', 'two']);
  __p(m[1]);
  __p(m.length);
}

void main() {
  __vybeMain();
  __check('one\n3');
}
