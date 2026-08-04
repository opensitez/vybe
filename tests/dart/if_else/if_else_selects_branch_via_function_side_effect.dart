// vybe-test: dart/if_else/if_else_selects_branch_via_function_side_effect
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
  var log = '';
  void record(String msg) { log = msg; }
  if (2 + 2 == 4) {
    record('math-ok');
  } else {
    record('math-fail');
  }
  __p(log);
}

void main() {
  __vybeMain();
  __check('math-ok');
}
