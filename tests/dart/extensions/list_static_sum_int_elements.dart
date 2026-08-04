// vybe-test: dart/extensions/list_static_sum_int_elements
// origin: languages/dart/tests/dart/test_extensions.rs

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

extension ListSum on List<int> {
  static int total(List<int> nums) {
    var sum = 0;
    for (var n in nums) {
      sum += n;
    }
    return sum;
  }
}
void __vybeMain() {
  __p(ListSum.total([1, 2, 3, 4]));
}

void main() {
  __vybeMain();
  __check('10');
}
