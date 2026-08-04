// vybe-test: dart/abstract_members/abstract_instance_used_in_list
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Node {
  int val();
}
class Leaf extends Node {
  int v;
  Leaf(this.v);
  int val() {
    return v;
  }
}
void __vybeMain() {
  List<Node> nodes = [Leaf(1), Leaf(2)];
  __p(nodes[0].val() + nodes[1].val());
}

void main() {
  __vybeMain();
  __check('3');
}
