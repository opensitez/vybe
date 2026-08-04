// vybe-test: dart/ternary_operator/ternary_null_check_with_value
// origin: languages/dart/tests/dart/test_ternary_operator.rs

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
  String? value = 'data';
  __p(value != null ? value : 'none');
}

void main() {
  __vybeMain();
  __check('data');
}
