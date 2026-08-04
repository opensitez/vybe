// vybe-test: dart/functions_advanced/required_result
// origin: languages/dart/tests/dart/test_functions_advanced.rs

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

void greet({required String name}) { __p('Hello $name'); }
void __vybeMain() { greet(name: 'Dart'); }

void main() {
  __vybeMain();
  __check('Hello Dart');
}
