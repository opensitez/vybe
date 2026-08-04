// vybe-test: dart/mixins_core/mixin_override_with_super_from_on_type
// origin: languages/dart/tests/dart/test_mixins_core.rs

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
  String val() {
    return 'b';
  }
}
mixin Mid on Base {
  String val() {
    return 'm';
  }
}
class End extends Base with Mid {}
void __vybeMain() {
  __p(End().val());
}

void main() {
  __vybeMain();
  __check('m');
}
