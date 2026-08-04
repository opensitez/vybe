// vybe-test: dart/null_operators/non_null_assert_in_return_expression
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

String? pick(String? a, String? b) => a ?? b!;
void __vybeMain() {
  __p(pick(null, 'second'));
}

void main() {
  __vybeMain();
  __check('second');
}
