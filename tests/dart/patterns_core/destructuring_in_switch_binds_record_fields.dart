// vybe-test: dart/patterns_core/destructuring_in_switch_binds_record_fields
// origin: languages/dart/tests/dart/test_patterns_core.rs

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
  var item = (id: 5, qty: 2);
  switch (item) {
    case (id: var i, qty: var q):
      __p(i * q);
      break;
    default:
      __p(0);
  }
}

void main() {
  __vybeMain();
  __check('10');
}
