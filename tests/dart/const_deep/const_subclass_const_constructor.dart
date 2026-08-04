// vybe-test: dart/const_deep/const_subclass_const_constructor
// origin: languages/dart/tests/dart/test_const_deep.rs

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
  final int n;
  const Base(this.n);
}
class Derived extends Base {
  const Derived(int v) : super(v);
}
void __vybeMain() {
  const d = Derived(11);
  __p(d.n);
}

void main() {
  __vybeMain();
  __check('11');
}
