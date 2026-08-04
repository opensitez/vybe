// vybe-test: dart/record_destructuring/destructure_mixed_after_reassignment
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
  var rec = (id: 1, tag: 'a');
  var (id: i, tag: t) = rec;
  rec = (id: 2, tag: 'b');
  __p(i);
  __p(t);
}

void main() {
  __vybeMain();
  __check('1\na');
}
