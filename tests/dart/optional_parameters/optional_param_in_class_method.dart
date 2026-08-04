// vybe-test: dart/optional_parameters/optional_param_in_class_method
// origin: languages/dart/tests/dart/test_optional_parameters.rs

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
  void say([String word = 'hi']) {
    __p(word);
  }
}
void __vybeMain() {
  Greeter().say();
  Greeter().say('yo');
}

void main() {
  __vybeMain();
  __check('hi\nyo');
}
