// vybe-test: dart/cascades/cascade_list_clear_then_add_rebuilds_contents
// origin: languages/dart/tests/dart/test_cascades.rs

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
  var nums = [9, 9, 9];
  nums..clear()..add(1)..add(2);
  __p(nums.isEmpty);
  __p(nums.join(','));
}

void main() {
  __vybeMain();
  __check('false\n1,2');
}
