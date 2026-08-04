// vybe-test: dart/set_algebra_deep/difference_with_self_yields_empty_set
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
  var s = {for (var i = 1; i <= 10; i++) i};
  var d = s.difference(s);
  __p(d.isEmpty);
  __p(d.length);
}

void main() {
  __vybeMain();
  __check('true\n0');
}
