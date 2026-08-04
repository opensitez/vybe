// vybe-test: dart/super_calls/super_constructor_with_this_shorthand_sibling
// origin: languages/dart/tests/dart/test_super_calls.rs

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
  int value;
  Node(this.value);
}
class Leaf extends Node {
  int tag;
  Leaf(this.tag, int v) : super(v);
}
void __vybeMain() {
  var leaf = Leaf(1, 8);
  __p(leaf.value);
}

void main() {
  __vybeMain();
  __check('8');
}
