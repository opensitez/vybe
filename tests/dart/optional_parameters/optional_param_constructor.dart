// vybe-test: dart/optional_parameters/optional_param_constructor
// origin: languages/dart/tests/dart/test_optional_parameters.rs

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

class Box {
  final int size;
  Box([this.size = 1]);
}
void __vybeMain() {
  __p(Box().size);
  __p(Box(9).size);
}

void main() {
  __vybeMain();
  __check('1\n9');
}
