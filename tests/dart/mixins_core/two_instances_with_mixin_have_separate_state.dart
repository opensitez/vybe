// vybe-test: dart/mixins_core/two_instances_with_mixin_have_separate_state
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

mixin State {
  int n = 0;
}
class Node with State {}
void __vybeMain() {
  var a = Node();
  var b = Node();
  a.n = 5;
  __p(b.n);
}

void main() {
  __vybeMain();
  __check('0');
}
