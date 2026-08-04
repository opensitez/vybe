// vybe-test: dart/list_sort_search/sort_strings_by_length_not_lexicographic
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
  var list = ['aaa', 'b', 'cc', 'd'];
  list.sort((a, b) => a.length.compareTo(b.length));
  __p(list.join(','));
}

void main() {
  __vybeMain();
  __check('b,d,cc,aaa');
}
