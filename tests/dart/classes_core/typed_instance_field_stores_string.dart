// vybe-test: dart/classes_core/typed_instance_field_stores_string
// origin: languages/dart/tests/dart/test_classes_core.rs

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

class Label {
  String text = 'hello';
}
void __vybeMain() {
  var l = Label();
  __p(l.text);
}

void main() {
  __vybeMain();
  __check('hello');
}
