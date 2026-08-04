// vybe-test: dart/collection_literals/list_collection_for_nested_in_expression
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
  var rows = [for (var r = 0; r < 2; r++) r];
  var flat = [for (var r in rows) for (var c = 0; c < 2; c++) r * 10 + c];
  __p(flat.join(','));
}

void main() {
  __vybeMain();
  __check('0,1,10,11');
}
