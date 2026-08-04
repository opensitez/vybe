// vybe-test: dart/switch_statements/string_switch_no_match_routes_to_default
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
  var fruit = 'cherry';
  switch (fruit) {
    case 'apple':
      __p('apple');
      break;
    case 'banana':
      __p('banana');
      break;
    default:
      __p('exotic');
  }
}

void main() {
  __vybeMain();
  __check('exotic');
}
