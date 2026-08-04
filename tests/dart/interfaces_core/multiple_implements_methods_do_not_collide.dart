// vybe-test: dart/interfaces_core/multiple_implements_methods_do_not_collide
// origin: languages/dart/tests/dart/test_interfaces_core.rs

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

abstract class Left {
  int left();
}
abstract class Right {
  int right();
}
class Both implements Left, Right {
  int left() {
    return 1;
  }
  int right() {
    return 10;
  }
}
void __vybeMain() {
  __p(Both().left() + Both().right());
}

void main() {
  __vybeMain();
  __check('11');
}
