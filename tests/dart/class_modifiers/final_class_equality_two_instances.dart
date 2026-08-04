// vybe-test: dart/class_modifiers/final_class_equality_two_instances
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

final class Item {
  int id;
  Item(this.id);
}
void __vybeMain() {
  __p(Item(1) == Item(1));
}

void main() {
  __vybeMain();
  __check('false');
}
