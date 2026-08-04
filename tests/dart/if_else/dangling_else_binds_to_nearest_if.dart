// vybe-test: dart/if_else/dangling_else_binds_to_nearest_if
// origin: languages/dart/tests/dart/test_if_else.rs

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
  if (true) {
    if (false) {
      __p('inner-then');
    } else {
      __p('inner-else');
    }
  }
}

void main() {
  __vybeMain();
  __check('inner-else');
}
