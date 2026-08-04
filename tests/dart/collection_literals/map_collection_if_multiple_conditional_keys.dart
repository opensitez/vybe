// vybe-test: dart/collection_literals/map_collection_if_multiple_conditional_keys
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
  var a = true;
  var b = true;
  var m = {if (a) 'x': 1, if (b) 'y': 2};
  __p(m.keys.join(','));
}

void main() {
  __vybeMain();
  __check('x,y');
}
