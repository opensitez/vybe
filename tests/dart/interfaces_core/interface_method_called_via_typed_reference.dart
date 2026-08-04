// vybe-test: dart/interfaces_core/interface_method_called_via_typed_reference
// origin: languages/dart/tests/dart/test_interfaces_core.rs

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

abstract class Op {
  int run(int n);
}
class Square implements Op {
  int run(int n) {
    return n * n;
  }
}
void __vybeMain() {
  Op o = Square();
  __p(o.run(5));
}

void main() {
  __vybeMain();
  __check('25');
}
