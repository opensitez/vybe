// vybe-test: dart/collections_advanced/list_fold_result
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

void __vybeMain() { var sum = [1, 2, 3, 4].fold(0, (acc, e) => acc + e); __p(sum); }

void main() {
  __vybeMain();
  __check('10');
}
