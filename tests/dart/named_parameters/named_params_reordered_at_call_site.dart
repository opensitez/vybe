// vybe-test: dart/named_parameters/named_params_reordered_at_call_site
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

void pair({int a = 0, int b = 0}) {
  __p('$a,$b');
}
void __vybeMain() {
  pair(b: 2, a: 1);
}

void main() {
  __vybeMain();
  __check('1,2');
}
