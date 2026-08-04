// vybe-test: dart/iterable_methods/iterable_element_at_on_map_entries
// origin: languages/dart/tests/dart/test_iterable_methods.rs

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
  var m = {1: 'one', 2: 'two'};
  var entries = m.entries;
  __p(entries.elementAt(0).key);
}

void main() {
  __vybeMain();
  __check('1');
}
