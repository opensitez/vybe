// vybe-test: dart/collection_literals/list_collection_if_nested_condition
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
  var outer = true;
  var inner = false;
  var list = [if (outer) if (inner) 1 else 2];
  __p(list[0]);
}

void main() {
  __vybeMain();
  __check('2');
}
