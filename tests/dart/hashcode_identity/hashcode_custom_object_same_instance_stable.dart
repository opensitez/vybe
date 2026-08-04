// vybe-test: dart/hashcode_identity/hashcode_custom_object_same_instance_stable
// origin: languages/dart/tests/dart/test_hashcode_identity.rs

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

class Widget {
  int id;
  Widget(this.id);
}
void __vybeMain() {
  var w = Widget(5);
  __p(w.hashCode == w.hashCode);
}

void main() {
  __vybeMain();
  __check('true');
}
