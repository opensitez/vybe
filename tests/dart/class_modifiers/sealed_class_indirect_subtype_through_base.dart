// vybe-test: dart/class_modifiers/sealed_class_indirect_subtype_through_base
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

sealed class Expr {}
class NumLit extends Expr {
  int n;
  NumLit(this.n);
}
class AddExpr extends Expr {
  Expr left;
  Expr right;
  AddExpr(this.left, this.right);
}
void __vybeMain() {
  var tree = AddExpr(NumLit(2), NumLit(3));
  __p((tree.left as NumLit).n);
}

void main() {
  __vybeMain();
  __check('2');
}
