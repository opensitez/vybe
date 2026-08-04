// vybe-test: dart/control_flow_advanced/while_accumulate
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
  var i = 1;
  var product = 1;
  while (i <= 5) {
    product *= i;
    i++;
  }
  __p(product);
}

void main() {
  __vybeMain();
  __check('120');
}
