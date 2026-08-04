// vybe-test: dart/mixin_linearization/mixin_on_with_super_in_getter
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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
  int get base {
    return 2;
  }
}
mixin Extra on Base {
  int get total {
    return super.base + 3;
  }
}
class Node extends Base with Extra {}
void __vybeMain() {
  __p(Node().total);
}

void main() {
  __vybeMain();
  __check('5');
}
