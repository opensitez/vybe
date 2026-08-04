// vybe-test: dart/switch_statements/switch_int_first_case_skips_later_unmatched_labels
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
  var n = 1;
  switch (n) {
    case 1:
      __p('hit-one');
      break;
    case 2:
      __p('hit-two');
      break;
    case 3:
      __p('hit-three');
      break;
  }
}

void main() {
  __vybeMain();
  __check('hit-one');
}
