// vybe-test: dart/map_entry_algorithms/entries_fold_product_of_values
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
  var m = {'a': 2, 'b': 3, 'c': 4};
  var product = m.entries.fold(1, (p, e) => p * e.value);
  __p(product);
}

void main() {
  __vybeMain();
  __check('24');
}
