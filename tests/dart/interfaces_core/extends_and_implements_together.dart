// vybe-test: dart/interfaces_core/extends_and_implements_together
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
  int baseVal() {
    return 1;
  }
}
abstract class Extra {
  int extraVal();
}
class Both extends Base implements Extra {
  int extraVal() {
    return 2;
  }
}
void __vybeMain() {
  var b = Both();
  __p(b.baseVal() + b.extraVal());
}

void main() {
  __vybeMain();
  __check('3');
}
