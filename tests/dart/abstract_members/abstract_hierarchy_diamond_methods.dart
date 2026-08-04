// vybe-test: dart/abstract_members/abstract_hierarchy_diamond_methods
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

abstract class Top {
  int top();
}
abstract class Left extends Top {
  int left();
}
class Bottom extends Left {
  int top() {
    return 1;
  }
  int left() {
    return 10;
  }
}
void __vybeMain() {
  __p(Bottom().top() + Bottom().left());
}

void main() {
  __vybeMain();
  __check('11');
}
