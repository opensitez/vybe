// vybe-test: dart/no_such_method/no_such_method_proxy_delegates_method_to_target
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

class Target {
  int doubleIt(int n) {
    return n * 2;
  }
}
class Proxy {
  final Target target;
  Proxy(this.target);
  @override
  dynamic noSuchMethod(Invocation inv) {
    if (inv.isMethod && inv.memberName == #doubleIt) {
      return target.doubleIt(inv.positionalArguments[0] as int);
    }
    return super.noSuchMethod(inv);
  }
}
void __vybeMain() {
  dynamic p = Proxy(Target());
  __p(p.doubleIt(5));
}

void main() {
  __vybeMain();
  __check('10');
}
