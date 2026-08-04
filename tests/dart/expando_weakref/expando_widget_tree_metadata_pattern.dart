// vybe-test: dart/expando_weakref/expando_widget_tree_metadata_pattern
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

class Element {
  int depth;
  Element(this.depth);
}
void __vybeMain() {
  final meta = Expando<String>();
  var root = Element(0);
  var child = Element(1);
  meta[root] = 'root';
  meta[child] = 'child';
  __p(meta[root]);
  __p(meta[child]);
}

void main() {
  __vybeMain();
  __check('root\nchild');
}
