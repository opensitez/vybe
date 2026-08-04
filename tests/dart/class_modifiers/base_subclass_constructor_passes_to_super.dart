// vybe-test: dart/class_modifiers/base_subclass_constructor_passes_to_super
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

base class Pair {
  int a;
  Pair(this.a);
}
class DoublePair extends Pair {
  DoublePair(int v) : super(v * 2);
}
void __vybeMain() {
  __p(DoublePair(3).a);
}

void main() {
  __vybeMain();
  __check('6');
}
