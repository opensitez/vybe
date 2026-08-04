// vybe-test: dart/no_such_method/no_such_method_forwarding_wrapper_pattern
// origin: languages/dart/tests/dart/test_no_such_method.rs

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

class Real {
  int add(int a, int b) {
    return a + b;
  }
}
class Wrapper {
  final Real inner;
  Wrapper(this.inner);
  @override
  dynamic noSuchMethod(Invocation inv) {
    if (inv.memberName == #add) {
      return inner.add(
        inv.positionalArguments[0] as int,
        inv.positionalArguments[1] as int,
      );
    }
    return null;
  }
}
void __vybeMain() {
  dynamic w = Wrapper(Real());
  __p(w.add(2, 3));
}

void main() {
  __vybeMain();
  __check('5');
}
