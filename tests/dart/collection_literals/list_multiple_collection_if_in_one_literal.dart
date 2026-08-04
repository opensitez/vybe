// vybe-test: dart/collection_literals/list_multiple_collection_if_in_one_literal
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
  var b = false;
  var list = [1, if (a) 2, if (b) 3, if (!b) 4];
  __p(list.join(','));
}

void main() {
  __vybeMain();
  __check('1,2,4');
}
