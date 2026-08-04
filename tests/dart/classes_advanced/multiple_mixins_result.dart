// vybe-test: dart/classes_advanced/multiple_mixins_result
// origin: languages/dart/tests/dart/test_classes_advanced.rs

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

mixin Greet { String greet() => 'hello'; }
mixin Bye { String bye() => 'goodbye'; }
class Person with Greet, Bye { String name; Person(this.name); }
void __vybeMain() {
  var p = Person('Alice');
  __p(p.greet());
}

void main() {
  __vybeMain();
  __check('hello');
}
