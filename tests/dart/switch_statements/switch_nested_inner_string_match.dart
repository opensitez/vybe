// vybe-test: dart/switch_statements/switch_nested_inner_string_match
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
  var outer = 1;
  var inner = 'b';
  switch (outer) {
    case 1:
      switch (inner) {
        case 'a':
          __p('inner-a');
          break;
        case 'b':
          __p('inner-b');
          break;
        default:
          __p('inner-other');
      }
      break;
    default:
      __p('outer-other');
  }
}

void main() {
  __vybeMain();
  __check('inner-b');
}
