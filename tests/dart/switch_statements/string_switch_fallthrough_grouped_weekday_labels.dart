// vybe-test: dart/switch_statements/string_switch_fallthrough_grouped_weekday_labels
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
  var day = 'Sat';
  switch (day) {
    case 'Sat':
    case 'Sun':
      __p('weekend');
      break;
    case 'Mon':
      __p('monday');
      break;
    default:
      __p('weekday');
  }
}

void main() {
  __vybeMain();
  __check('weekend');
}
