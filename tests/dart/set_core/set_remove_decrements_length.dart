// vybe-test: dart/set_core/set_remove_decrements_length
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
  var s = {10, 20, 30};
  s.remove(20);
  __p(s.length);
  __p(s.contains(20));
}

void main() {
  __vybeMain();
  __check('2\nfalse');
}
