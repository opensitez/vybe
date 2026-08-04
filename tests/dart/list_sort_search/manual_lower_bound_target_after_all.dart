// vybe-test: dart/list_sort_search/manual_lower_bound_target_after_all
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
  var list = [2, 4, 6];
  var target = 10;
  var lo = 0;
  var hi = list.length;
  while (lo < hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid; }
  }
  __p(lo);
}

void main() {
  __vybeMain();
  __check('3');
}
