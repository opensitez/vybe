// vybe-test: dart/functions_core/local_function_shadows_outer_name
// origin: languages/dart/tests/dart/test_functions_core.rs

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

var value = 1;
int getValue() {
  return value;
}
void __vybeMain() {
  int getValue() {
    return 99;
  }
  __p(getValue());
}

void main() {
  __vybeMain();
  __check('99');
}
