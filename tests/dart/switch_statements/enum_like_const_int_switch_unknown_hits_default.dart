// vybe-test: dart/switch_statements/enum_like_const_int_switch_unknown_hits_default
// origin: languages/dart/tests/dart/test_switch_statements.rs

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
  const red = 0;
  const green = 1;
  const blue = 2;
  var color = 9;
  switch (color) {
    case red:
      __p('red');
      break;
    case green:
      __p('green');
      break;
    case blue:
      __p('blue');
      break;
    default:
      __p('unknown');
  }
}

void main() {
  __vybeMain();
  __check('unknown');
}
