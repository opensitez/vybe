// vybe-test: dart/switch_statements/switch_break_inside_switch_does_not_continue_outer_loop
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
  for (var i = 0; i < 2; i++) {
    switch (i) {
      case 0:
        __p('loop-0');
        break;
      case 1:
        __p('loop-1');
        break;
    }
    __p('after-$i');
  }
}

void main() {
  __vybeMain();
  __check('loop-0\nafter-0\nloop-1\nafter-1');
}
