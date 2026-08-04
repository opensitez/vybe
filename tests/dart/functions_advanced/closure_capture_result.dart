// vybe-test: dart/functions_advanced/closure_capture_result
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

void __vybeMain() { var x = 10; var fn = () => x * 2; __p(fn()); }

void main() {
  __vybeMain();
  __check('20');
}
