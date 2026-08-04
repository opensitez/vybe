// vybe-test: dart/functions_core/return_from_nested_block_inside_loop
// origin: languages/dart/tests/dart/test_functions_core.rs

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

int firstNegative(List<int> nums) {
  for (var n in nums) {
    if (n < 0) {
      return n;
    }
  }
  return 0;
}
void __vybeMain() {
  __p(firstNegative([1, 2, -3, 4]));
}

void main() {
  __vybeMain();
  __check('-3');
}
