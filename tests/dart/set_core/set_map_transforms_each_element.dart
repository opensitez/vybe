// vybe-test: dart/set_core/set_map_transforms_each_element
// origin: languages/dart/tests/dart/test_set_core.rs

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
  var s = {1, 2, 3};
  var doubled = s.map((e) => e * 2).toList()..sort();
  __p(doubled.join(','));
}

void main() {
  __vybeMain();
  __check('2,4,6');
}
