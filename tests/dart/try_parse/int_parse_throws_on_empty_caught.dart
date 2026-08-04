// vybe-test: dart/try_parse/int_parse_throws_on_empty_caught
// origin: languages/dart/tests/dart/test_try_parse.rs

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

void __vybeMain() {
  try {
    int.parse('');
  } catch (e) {
    __p('caught');
  }
}

void main() {
  __vybeMain();
  __check('caught');
}
