// vybe-test: dart/class_modifiers/sealed_class_field_on_subtype
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

sealed class Event {}
class Click extends Event {
  int x;
  int y;
  Click(this.x, this.y);
}
void __vybeMain() {
  var e = Click(3, 4);
  __p(e.x + e.y);
}

void main() {
  __vybeMain();
  __check('7');
}
