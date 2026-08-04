// vybe-test: dart/factory_constructors_deep/factory_redirect_chain_to_primary
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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
  int depth;
  Node(this.depth);
  Node.root() : depth = 0;
  factory Node.zero() = Node.root;
}
void __vybeMain() {
  __p(Node.zero().depth);
}

void main() {
  __vybeMain();
  __check('0');
}
