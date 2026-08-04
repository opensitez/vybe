// vybe-test: dart/control_flow_advanced/switch_int
// origin: languages/dart/tests/dart/test_control_flow_advanced.rs

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
  var n = 2;
  switch (n) {
    case 1: __p('one'); break;
    case 2: __p('two'); break;
    case 3: __p('three'); break;
    default: __p('many');
  }
}

void main() {
  __vybeMain();
  __check('two');
}
