// vybe-test: dart/records_core/record_field_read_after_variable_assignment
// origin: languages/dart/tests/dart/test_records_core.rs

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
  var a = (10, 20);
  var b = a;
  __p(b.$1);
  __p(b.$2);
}

void main() {
  __vybeMain();
  __check('10\n20');
}
