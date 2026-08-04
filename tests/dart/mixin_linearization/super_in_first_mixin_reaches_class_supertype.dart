// vybe-test: dart/mixin_linearization/super_in_first_mixin_reaches_class_supertype
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
  String root() {
    return 'base';
  }
}
mixin Wrap on Base {
  String root() {
    return super.root() + '-wrap';
  }
}
class Node extends Base with Wrap {}
void __vybeMain() {
  __p(Node().root());
}

void main() {
  __vybeMain();
  __check('base-wrap');
}
