// vybe-test: dart/enums_advanced/enum_in_if
// origin: languages/dart/tests/dart/test_enums_advanced.rs

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

enum Direction { up, down, left, right }
void __vybeMain() {
  var d = Direction.up;
  if (d == Direction.up) { __p('going up'); }
}

void main() {
  __vybeMain();
  __check('going up');
}
