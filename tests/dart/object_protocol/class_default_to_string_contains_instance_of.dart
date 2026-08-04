// vybe-test: dart/object_protocol/class_default_to_string_contains_instance_of
// origin: languages/dart/tests/dart/test_object_protocol.rs

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

class Widget {}
void __vybeMain() {
  __p(Widget().toString().contains('Widget'));
}

void main() {
  __vybeMain();
  __check('true');
}
