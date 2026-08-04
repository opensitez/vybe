// vybe-test: dart/collection_literals/map_collection_for_string_keys_from_numbers
// origin: languages/dart/tests/dart/test_collection_literals.rs

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
  var m = {for (var i = 1; i <= 2; i++) 'k$i': i};
  __p(m['k1']);
  __p(m['k2']);
}

void main() {
  __vybeMain();
  __check('1\n2');
}
