// vybe-test: dart/closures/nested_closure_reads_two_outer_variables
// origin: languages/dart/tests/dart/test_closures.rs

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
  var a = 2;
  var b = 3;
  var compute = () {
    var inner = () => a * b;
    return inner();
  };
  __p(compute());
}

void main() {
  __vybeMain();
  __check('6');
}
