// vybe-test: dart/mixins_core/basic_mixin_method_available_on_class
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

mixin Greet {
  String hello() {
    return 'hi';
  }
}
class Person with Greet {}
void __vybeMain() {
  __p(Person().hello());
}

void main() {
  __vybeMain();
  __check('hi');
}
