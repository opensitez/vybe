// vybe-test: dart/covariant_keyword/covariant_param_with_super_call
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

class Base {
  int n;
  Base(this.n);
}
class Derived extends Base {
  Derived(int v) : super(v);
}
class Processor {
  void run(Base b) {}
}
class DerivedProcessor extends Processor {
  @override
  void run(covariant Derived d) {
    __p(d.n);
  }
}
void __vybeMain() {
  DerivedProcessor().run(Derived(6));
}

void main() {
  __vybeMain();
  __check('6');
}
