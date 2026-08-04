// vybe-test: dart/named_parameters/three_named_params_any_order
// origin: languages/dart/tests/dart/test_named_parameters.rs

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

void triple({int x = 0, int y = 0, int z = 0}) {
  __p('$x$y$z');
}
void __vybeMain() {
  triple(z: 3, x: 1, y: 2);
}

void main() {
  __vybeMain();
  __check('123');
}
