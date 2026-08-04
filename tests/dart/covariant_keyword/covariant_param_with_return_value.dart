// vybe-test: dart/covariant_keyword/covariant_param_with_return_value
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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

class Node {}
class Leaf extends Node {}
class Tree {
  String tag(Node n) {
    return 'node';
  }
}
class LeafTree extends Tree {
  @override
  String tag(covariant Leaf l) {
    return 'leaf';
  }
}
void __vybeMain() {
  __p(LeafTree().tag(Leaf()));
}

void main() {
  __vybeMain();
  __check('leaf');
}
