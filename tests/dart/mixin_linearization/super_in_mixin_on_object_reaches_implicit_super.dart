// vybe-test: dart/mixin_linearization/super_in_mixin_on_object_reaches_implicit_super
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

mixin M on Object {
  String describe() {
    return 'mixin';
  }
}
class T with M {}
void __vybeMain() {
  __p(T().describe());
}

void main() {
  __vybeMain();
  __check('mixin');
}
