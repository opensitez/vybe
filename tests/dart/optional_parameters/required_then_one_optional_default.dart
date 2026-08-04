// vybe-test: dart/optional_parameters/required_then_one_optional_default
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

void log(String level, [String detail = 'none']) {
  __p('$level:$detail');
}
void __vybeMain() {
  log('INFO');
}

void main() {
  __vybeMain();
  __check('INFO:none');
}
