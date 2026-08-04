// vybe-test: dart/switch_statements/switch_inside_for_loop_selects_per_iteration
// origin: languages/dart/tests/dart/test_switch_statements.rs

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
  for (var i = 0; i < 3; i++) {
    switch (i) {
      case 0:
        __p('zero');
        break;
      case 1:
        __p('one');
        break;
      case 2:
        __p('two');
        break;
    }
  }
}

void main() {
  __vybeMain();
  __check('zero\none\ntwo');
}
