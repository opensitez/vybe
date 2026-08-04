// vybe-test: dart/list_sort_search/sort_then_lower_bound_matches_index_of_for_unique
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
  var list = [9, 3, 7, 1, 5];
  list.sort();
  var target = 7;
  var idx = list.indexOf(target);
  var lo = 0;
  var hi = list.length;
  while (lo < hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid; }
  }
  __p(idx);
  __p(lo);
  __p(list[idx]);
}

void main() {
  __vybeMain();
  __check('3\n3\n7');
}
