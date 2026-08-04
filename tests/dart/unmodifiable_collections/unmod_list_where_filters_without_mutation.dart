// vybe-test: dart/unmodifiable_collections/unmod_list_where_filters_without_mutation
// origin: languages/dart/tests/dart/test_unmodifiable_collections.rs

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
  var frozen = List.unmodifiable([1, 2, 3, 4]);
  __p(frozen.where((n) => n.isEven).join(','));
}

void main() {
  __vybeMain();
  __check('2,4');
}
