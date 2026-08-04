// vybe-test: dart/interfaces_core/implements_different_from_extends
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

class Base {
  int baseOnly() {
    return 1;
  }
}
abstract class Port {
  int portVal();
}
class Svc extends Base implements Port {
  int portVal() {
    return 2;
  }
}
void __vybeMain() {
  var s = Svc();
  __p(s.baseOnly() + s.portVal());
}

void main() {
  __vybeMain();
  __check('3');
}
