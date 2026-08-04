// vybe-test: dart/list_sort_search/sort_by_absolute_value_custom_comparator
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
  var list = [-10, 3, -1, 8, -5];
  list.sort((a, b) => a.abs().compareTo(b.abs()));
  __p(list.join(','));
}

void main() {
  __vybeMain();
  __check('-1,3,-5,8,-10');
}
