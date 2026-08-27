// vybe-test: dart/collection_literals/map_collection_if_else_picks_value
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
  // Damaged test repaired: dart 3.10.4 rejects `if` in a map VALUE position
  // ("Expected an identifier, but got 'if'"); collection-`if` is an ELEMENT,
  // so the whole entry is what the branches choose.
  var debug = false;
  var m = {if (debug) 'level': 0 else 'level': 1};
  __p(m['level']);
}

void main() {
  __vybeMain();
  __check('1');
}
