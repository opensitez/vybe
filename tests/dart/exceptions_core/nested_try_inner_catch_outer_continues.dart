// vybe-test: dart/exceptions_core/nested_try_inner_catch_outer_continues
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
    try {
      throw 'inner';
    } catch (e) {
      __p('inner:$e');
    }
    __p('middle');
  } catch (e) {
    __p('outer');
  }
}

void main() {
  __vybeMain();
  __check('inner:inner\nmiddle');
}
