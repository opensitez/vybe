// vybe-test: dart/map_core/map_map_transforms_entries_into_new_map
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
  var m = {'a': 1, 'b': 2};
  var doubled = m.map((k, v) => MapEntry(k, v * 2));
  __p(doubled['a']);
  __p(doubled['b']);
}

void main() {
  __vybeMain();
  __check('2\n4');
}
