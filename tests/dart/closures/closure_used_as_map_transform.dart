// vybe-test: dart/closures/closure_used_as_map_transform
// origin: languages/dart/tests/dart/test_closures.rs

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
  var offset = 10;
  var nums = [1, 2, 3];
  var mapped = nums.map((n) => n + offset).toList();
  __p(mapped.join('-'));
}

void main() {
  __vybeMain();
  __check('11-12-13');
}
