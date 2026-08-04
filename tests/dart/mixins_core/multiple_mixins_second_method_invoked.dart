// vybe-test: dart/mixins_core/multiple_mixins_second_method_invoked
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

mixin Fly {
  String mode() {
    return 'air';
  }
}
mixin Swim {
  String mode() {
    return 'water';
  }
}
class Duck with Fly, Swim {}
void __vybeMain() {
  __p(Duck().mode());
}

void main() {
  __vybeMain();
  __check('water');
}
