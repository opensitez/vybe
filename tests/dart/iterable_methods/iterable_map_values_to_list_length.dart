// vybe-test: dart/iterable_methods/iterable_map_values_to_list_length
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
  var m = {'x': 10, 'y': 20, 'z': 30};
  __p(m.values.toList().length);
  __p(m.values.elementAt(1));
}

void main() {
  __vybeMain();
  __check('3\n20');
}
