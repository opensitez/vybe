// vybe-test: dart/ternary_operator/ternary_nested_lowest_branch
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
  var score = 60;
  var grade = score >= 90 ? 'A' : score >= 80 ? 'B' : 'C';
  __p(grade);
}

void main() {
  __vybeMain();
  __check('C');
}
