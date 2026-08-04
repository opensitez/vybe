// vybe-test: dart/class_modifiers/sealed_switch_nested_in_function
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

sealed class Option {}
class Some extends Option {
  int value;
  Some(this.value);
}
class None extends Option {}
int unwrap(Option o) {
  switch (o) {
    case Some(value: var v):
      return v;
    case None():
      return -1;
  }
}
void __vybeMain() {
  __p(unwrap(Some(99)));
}

void main() {
  __vybeMain();
  __check('99');
}
