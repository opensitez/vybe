// vybe-test: dart/collection_literals/map_collection_for_in_with_string_values
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
  var keys = ['x', 'y'];
  var m = {for (var k in keys) k: k.length};
  __p(m['x']);
  __p(m['y']);
}

void main() {
  __vybeMain();
  __check('1\n1');
}
