// vybe-test: dart/function_apply/apply_instance_method_with_named
// origin: languages/dart/tests/dart/test_function_apply.rs

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
  String greet(String name, {String title = 'Mr'}) {
    return '$title $name';
  }
}
void __vybeMain() {
  var g = Greeter();
  __p(Function.apply(g.greet, ['Lee'], {#title: 'Dr'}));
}

void main() {
  __vybeMain();
  __check('Dr Lee');
}
