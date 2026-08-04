// vybe-test: dart/list_sort_search/sort_then_index_of_finds_first_duplicate
// origin: languages/dart/tests/dart/test_list_sort_search.rs

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
  var list = [2, 1, 2, 3];
  list.sort();
  __p(list.indexOf(2));
  __p(list.lastIndexOf(2));
}

void main() {
  __vybeMain();
  __check('0\n1');
}
