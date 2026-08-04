// vybe-test: dart/super_calls/super_in_mixin_on_class_hierarchy
// origin: languages/dart/tests/dart/test_super_calls.rs

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

class Root {
  int rootVal() {
    return 2;
  }
}
class Mid extends Root {
  int rootVal() {
    return super.rootVal() + 3;
  }
}
void __vybeMain() {
  __p(Mid().rootVal());
}

void main() {
  __vybeMain();
  __check('5');
}
