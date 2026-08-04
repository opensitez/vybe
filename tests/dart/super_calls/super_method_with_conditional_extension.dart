// vybe-test: dart/super_calls/super_method_with_conditional_extension
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

class Base {
  int score(bool bonus) {
    return bonus ? 2 : 1;
  }
}
class Plus extends Base {
  int score(bool bonus) {
    return super.score(bonus) + 10;
  }
}
void __vybeMain() {
  __p(Plus().score(true));
}

void main() {
  __vybeMain();
  __check('12');
}
