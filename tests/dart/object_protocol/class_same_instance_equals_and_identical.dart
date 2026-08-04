// vybe-test: dart/object_protocol/class_same_instance_equals_and_identical
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

class Item {
  int id;
  Item(this.id);
}
void __vybeMain() {
  var x = Item(5);
  __p(x == x);
  __p(identical(x, x));
}

void main() {
  __vybeMain();
  __check('true\ntrue');
}
