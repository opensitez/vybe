// vybe-test: dart/list_core/list_sort_with_custom_comparator_descending
// origin: languages/dart/tests/dart/test_list_core.rs

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
  var list = [3, 1, 2];
  list.sort((a, b) => b.compareTo(a));
  __p(list.join(','));
}

void main() {
  __vybeMain();
  __check('3,2,1');
}
