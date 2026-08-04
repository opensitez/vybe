// vybe-test: dart/mixins_core/with_single_mixin_on_class_with_own_field
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

mixin Named {
  String label() {
    return 'named';
  }
}
class Widget {
  int id = 1;
  Widget();
}
class NamedWidget extends Widget with Named {}
void __vybeMain() {
  var w = NamedWidget();
  __p(w.label());
}

void main() {
  __vybeMain();
  __check('named');
}
