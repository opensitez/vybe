// vybe-test: dart/classes_core/super_constructor_passes_argument_to_base
// origin: languages/dart/tests/dart/test_classes_core.rs

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

class Base {
  int n;
  Base(this.n);
}
class Derived extends Base {
  Derived(int v) : super(v);
}
void __vybeMain() {
  __p(Derived(9).n);
}

void main() {
  __vybeMain();
  __check('9');
}
