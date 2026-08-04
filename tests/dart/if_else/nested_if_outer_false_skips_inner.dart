// vybe-test: dart/if_else/nested_if_outer_false_skips_inner
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
  var x = -1;
  var y = 10;
  if (x > 0) {
    if (y > 0) {
      __p('inner');
    }
  } else {
    __p('outer-else');
  }
}

void main() {
  __vybeMain();
  __check('outer-else');
}
