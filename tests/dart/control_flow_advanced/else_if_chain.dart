// vybe-test: dart/control_flow_advanced/else_if_chain
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
  var score = 75;
  if (score >= 90) {
    __p('A');
  } else if (score >= 80) {
    __p('B');
  } else if (score >= 70) {
    __p('C');
  } else {
    __p('F');
  }
}

void main() {
  __vybeMain();
  __check('C');
}
