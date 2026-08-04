// vybe-test: dart/list_sort_search/sort_then_index_of_on_last_element
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
  var list = [1, 5, 3, 4, 2];
  list.sort();
  __p(list.indexOf(5));
  __p(list.last);
}

void main() {
  __vybeMain();
  __check('4\n5');
}
