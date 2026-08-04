// vybe-test: dart/sync_star_generators/sync_star_tree_preorder_traversal
// origin: languages/dart/tests/dart/test_sync_star_generators.rs

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

Iterable<int> preorder(List<int> nodes) sync* {
  if (nodes.isEmpty) return;
  yield nodes[0];
  if (nodes.length > 1) { yield* preorder(nodes.sublist(1)); }
}
void __vybeMain() {
  __p(preorder([1, 2, 3]).join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
