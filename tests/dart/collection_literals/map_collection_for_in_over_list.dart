// vybe-test: dart/collection_literals/map_collection_for_in_over_list
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
  var names = ['a', 'b'];
  var m = {for (var i = 0; i < names.length; i++) names[i]: i};
  __p(m['a']);
  __p(m['b']);
}

void main() {
  __vybeMain();
  __check('0\n1');
}
