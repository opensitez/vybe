// vybe-test: dart/control_flow_advanced/break_result
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
  var sum = 0;
  for (var i = 1; i <= 10; i++) {
    if (i > 5) break;
    sum += i;
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('15');
}
