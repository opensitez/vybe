// vybe-test: dart/operator_overloading/operator_equals_null_object_false
// origin: languages/dart/tests/dart/test_operator_overloading.rs

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

class Node {
  int id;
  Node(this.id);
  bool operator ==(Object other) {
    if (other == null) return false;
    return other is Node && id == other.id;
  }
}
void __vybeMain() {
  __p(Node(1) == null);
}

void main() {
  __vybeMain();
  __check('false');
}
