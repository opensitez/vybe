// vybe-test: dart/list_sort_search/sort_then_manual_search_missing_target
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
  var list = [1, 3, 5, 7];
  list.sort();
  var target = 4;
  var found = -1;
  var lo = 0;
  var hi = list.length - 1;
  while (lo <= hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] == target) { found = mid; break; }
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid - 1; }
  }
  __p(found);
}

void main() {
  __vybeMain();
  __check('-1');
}
