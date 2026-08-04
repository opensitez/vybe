// vybe-test: dart/exceptions_core/try_catch_preserves_code_after_block
// origin: languages/dart/tests/dart/test_exceptions_core.rs

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
    throw 'stop';
  } catch (e) {
    __p('caught');
  }
  __p('continued');
}

void main() {
  __vybeMain();
  __check('caught\ncontinued');
}
