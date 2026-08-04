// vybe-test: dart/class_modifiers/sealed_switch_returns_value_from_second_subtype
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

sealed class Color {}
class Red extends Color {}
class Blue extends Color {}
String label(Color c) {
  switch (c) {
    case Red():
      return 'r';
    case Blue():
      return 'b';
  }
}
void __vybeMain() {
  __p(label(Blue()));
}

void main() {
  __vybeMain();
  __check('b');
}
