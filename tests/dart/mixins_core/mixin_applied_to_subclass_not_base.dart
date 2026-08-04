// vybe-test: dart/mixins_core/mixin_applied_to_subclass_not_base
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
  String baseId() {
    return 'b';
  }
}
class Mid extends Base {}
mixin Tag {
  String tagId() {
    return 't';
  }
}
class Leaf extends Mid with Tag {}
void __vybeMain() {
  var l = Leaf();
  __p(l.baseId() + l.tagId());
}

void main() {
  __vybeMain();
  __check('bt');
}
