// vybe-test: dart/mixins_core/mixin_with_extends_superclass_method_still_visible
// origin: languages/dart/tests/dart/test_mixins_core.rs

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

class Root {
  int rootVal() {
    return 10;
  }
}
mixin Branch {
  int branchVal() {
    return 1;
  }
}
class Tree extends Root with Branch {}
void __vybeMain() {
  var t = Tree();
  __p(t.rootVal() + t.branchVal());
}

void main() {
  __vybeMain();
  __check('11');
}
