// vybe-test: dart/set_core/set_union_with_empty_equals_original
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
  var a = {1, 2};
  var c = a.union(<int>{});
  __p(c.length);
  __p(c.contains(1));
  __p(c.contains(2));
}

void main() {
  __vybeMain();
  __check('2\ntrue\ntrue');
}
