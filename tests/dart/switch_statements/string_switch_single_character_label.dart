// vybe-test: dart/switch_statements/string_switch_single_character_label
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
  var ch = 'z';
  switch (ch) {
    case 'a':
      __p('vowel-start');
      break;
    case 'z':
      __p('last-letter');
      break;
    default:
      __p('middle');
  }
}

void main() {
  __vybeMain();
  __check('last-letter');
}
