// vybe-test: dart/record_destructuring/destructure_named_record_in_for_in
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
  var users = [(name: 'Ann', score: 10), (name: 'Bob', score: 20)];
  var total = 0;
  for (var (name: _, score: s) in users) {
    total += s;
  }
  __p(total);
}

void main() {
  __vybeMain();
  __check('30');
}
