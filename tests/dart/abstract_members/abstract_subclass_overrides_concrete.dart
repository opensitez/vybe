// vybe-test: dart/abstract_members/abstract_subclass_overrides_concrete
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Base {
  String tag() {
    return 'base';
  }
  String label();
}
class Sub extends Base {
  String label() {
    return super.tag() + '-sub';
  }
}
void __vybeMain() {
  __p(Sub().label());
}

void main() {
  __vybeMain();
  __check('base-sub');
}
