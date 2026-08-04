// vybe-test: dart/record_destructuring/destructure_positional_in_nested_loop
// origin: languages/dart/tests/dart/test_record_destructuring.rs

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
  var grid = [[(1, 1), (2, 2)], [(3, 3), (4, 4)]];
  var sum = 0;
  for (var row in grid) {
    for (var (x, y) in row) {
      sum += x + y;
    }
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('20');
}
