// vybe-test: dart/unmodifiable_collections/unmod_list_index_of_finds_element
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
  var frozen = List.unmodifiable([10, 20, 30]);
  __p(frozen.indexOf(20));
  __p(frozen.lastIndexOf(20));
}

void main() {
  __vybeMain();
  __check('1\n1');
}
