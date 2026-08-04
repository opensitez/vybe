// vybe-test: dart/typedefs_core/generic_typedef_comparator_sorts_descending
// origin: languages/dart/tests/dart/test_typedefs_core.rs

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

typedef IntList = List<int>;
typedef CompareFn = int Function(int, int);
void sortDesc(IntList items, CompareFn cmp) {
  items.sort(cmp);
}
void __vybeMain() {
  IntList nums = [1, 3, 2];
  sortDesc(nums, (a, b) => b.compareTo(a));
  __p(nums.join(','));
}

void main() {
  __vybeMain();
  __check('3,2,1');
}
