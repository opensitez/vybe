// vybe-test: dart/collection_literals/list_collection_if_false_omits_element
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
  var show = false;
  var list = [1, 2, if (show) 3];
  __p(list.length);
  __p(list.join(','));
}

void main() {
  __vybeMain();
  __check('2\n1,2');
}
