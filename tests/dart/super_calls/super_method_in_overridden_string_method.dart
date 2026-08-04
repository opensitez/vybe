// vybe-test: dart/super_calls/super_method_in_overridden_string_method
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

class Greeter {
  String hello() {
    return 'hi';
  }
}
class Loud extends Greeter {
  String hello() {
    return super.hello() + '!';
  }
}
void __vybeMain() {
  __p(Loud().hello());
}

void main() {
  __vybeMain();
  __check('hi!');
}
