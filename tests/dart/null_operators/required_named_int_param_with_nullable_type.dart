// vybe-test: dart/null_operators/required_named_int_param_with_nullable_type
// origin: languages/dart/tests/dart/test_null_operators.rs

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

void show({required int? count}) {
  __p(count ?? 0);
}
void __vybeMain() {
  show(count: null);
}

void main() {
  __vybeMain();
  __check('0');
}
