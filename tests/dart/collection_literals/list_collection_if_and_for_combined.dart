// vybe-test: dart/collection_literals/list_collection_if_and_for_combined
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
  var includeExtra = true;
  var list = [for (var i = 1; i <= 3; i++) i, if (includeExtra) 99];
  __p(list.join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3,99');
}
