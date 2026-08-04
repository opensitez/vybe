// vybe-test: dart/functions_core/arrow_function_block_body_with_multiple_statements
// origin: languages/dart/tests/dart/test_functions_core.rs

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
  var describe = (int n) {
    var sign = n < 0 ? 'neg' : 'pos';
    return '$sign:$n';
  };
  __p(describe(-4));
  __p(describe(4));
}

void main() {
  __vybeMain();
  __check('neg:-4\npos:4');
}
