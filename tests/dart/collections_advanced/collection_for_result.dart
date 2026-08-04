// vybe-test: dart/collections_advanced/collection_for_result
// origin: languages/dart/tests/dart/test_collections_advanced.rs

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

void __vybeMain() { var list = [for (var i = 0; i < 3; i++) i]; __p(list.length); }

void main() {
  __vybeMain();
  __check('3');
}
