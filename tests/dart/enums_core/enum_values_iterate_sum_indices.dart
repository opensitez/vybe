// vybe-test: dart/enums_core/enum_values_iterate_sum_indices
// origin: languages/dart/tests/dart/test_enums_core.rs

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

enum Digit { zero, one, two }
void __vybeMain() {
  var sum = 0;
  for (var d in Digit.values) {
    sum += d.index;
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('3');
}
