// vybe-test: dart/records_core/function_returns_positional_record
// origin: languages/dart/tests/dart/test_records_core.rs

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

({int, int}) minMax(List<int> nums) {
  return (nums.first, nums.last);
}
void __vybeMain() {
  var result = minMax([3, 9, 1, 7]);
  __p(result.$1);
  __p(result.$2);
}

void main() {
  __vybeMain();
  __check('3\n7');
}
