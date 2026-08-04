// vybe-test: dart/expando_weakref/expando_on_two_custom_instances
// origin: languages/dart/tests/dart/test_expando_weakref.rs

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
  String label;
  Node(this.label);
}
void __vybeMain() {
  final bag = Expando<int>();
  var n1 = Node('a');
  var n2 = Node('b');
  bag[n1] = 1;
  bag[n2] = 2;
  __p(bag[n1]);
  __p(bag[n2]);
}

void main() {
  __vybeMain();
  __check('1\n2');
}
