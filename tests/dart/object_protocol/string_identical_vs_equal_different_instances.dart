// vybe-test: dart/object_protocol/string_identical_vs_equal_different_instances
// origin: languages/dart/tests/dart/test_object_protocol.rs

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
  var a = String.fromCharCode(65);
  var b = String.fromCharCode(65);
  __p(identical(a, b));
  __p(a == b);
}

void main() {
  __vybeMain();
  __check('false\ntrue');
}
