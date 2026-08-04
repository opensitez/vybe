// vybe-test: dart/records_core/nested_record_field_access
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
  var outer = ((1, 2), label: 'pair');
  __p(outer.$1.$1);
  __p(outer.$1.$2);
  __p(outer.label);
}

void main() {
  __vybeMain();
  __check('1\n2\npair');
}
