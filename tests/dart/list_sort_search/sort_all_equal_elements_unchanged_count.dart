// vybe-test: dart/list_sort_search/sort_all_equal_elements_unchanged_count
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
  var list = [7, 7, 7, 7];
  list.sort();
  __p(list.join(','));
  __p(list.length);
}

void main() {
  __vybeMain();
  __check('7,7,7,7\n4');
}
