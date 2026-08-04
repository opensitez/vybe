// vybe-test: dart/mixins_core/mixin_field_accessible_on_host_class
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

mixin Tagged {
  String tag = 'm';
}
class Item with Tagged {}
void __vybeMain() {
  __p(Item().tag);
}

void main() {
  __vybeMain();
  __check('m');
}
