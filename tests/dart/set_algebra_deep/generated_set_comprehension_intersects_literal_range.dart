// vybe-test: dart/set_algebra_deep/generated_set_comprehension_intersects_literal_range
// origin: languages/dart/tests/dart/test_set_algebra_deep.rs

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
  var generated = {for (var i = 0; i < 10; i++) i * 2};
  var keep = {4, 6, 8, 9};
  var r = generated.intersection(keep).toList()..sort();
  __p(r.length);
  __p(r.join(','));
}

void main() {
  __vybeMain();
  __check('3\n4,6,8');
}
