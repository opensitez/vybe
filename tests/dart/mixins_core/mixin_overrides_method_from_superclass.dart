// vybe-test: dart/mixins_core/mixin_overrides_method_from_superclass
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
  String talk() {
    return 'base';
  }
}
mixin Loud on Base {
  String talk() {
    return 'loud';
  }
}
class Speaker extends Base with Loud {}
void __vybeMain() {
  __p(Speaker().talk());
}

void main() {
  __vybeMain();
  __check('loud');
}
