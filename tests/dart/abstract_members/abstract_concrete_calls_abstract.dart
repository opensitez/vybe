// vybe-test: dart/abstract_members/abstract_concrete_calls_abstract
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

abstract class Template {
  String step1() {
    return 'a';
  }
  String step2();
  String run() {
    return step1() + step2();
  }
}
class Impl extends Template {
  String step2() {
    return 'b';
  }
}
void __vybeMain() {
  __p(Impl().run());
}

void main() {
  __vybeMain();
  __check('ab');
}
