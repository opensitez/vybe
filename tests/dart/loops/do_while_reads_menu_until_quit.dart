// vybe-test: dart/loops/do_while_reads_menu_until_quit
// origin: languages/dart/tests/dart/test_loops.rs

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
  var choices = [1, 2, 0];
  var idx = 0;
  var picks = 0;
  do {
    picks++;
    idx++;
  } while (choices[idx - 1] != 0 && idx < choices.length);
  __p(picks);
}

void main() {
  __vybeMain();
  __check('3');
}
