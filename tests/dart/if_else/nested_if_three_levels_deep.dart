// vybe-test: dart/if_else/nested_if_three_levels_deep
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
  var a = 1;
  var b = 2;
  var c = 3;
  if (a == 1) {
    if (b == 2) {
      if (c == 3) {
        __p('deep');
      }
    }
  }
}

void main() {
  __vybeMain();
  __check('deep');
}
