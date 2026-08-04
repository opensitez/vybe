// vybe-test: dart/super_calls/super_constructor_with_string_arg
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

class Named {
  String name;
  Named(this.name);
}
class Alias extends Named {
  Alias(String n) : super(n);
}
void __vybeMain() {
  __p(Alias('dart').name);
}

void main() {
  __vybeMain();
  __check('dart');
}
