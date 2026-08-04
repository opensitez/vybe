// vybe-test: dart/hashcode_identity/identical_two_distinct_string_instances
// origin: languages/dart/tests/dart/test_hashcode_identity.rs

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
  var a = String.fromCharCodes([97, 98]);
  var b = String.fromCharCodes([97, 98]);
  __p(identical(a, b));
  __p(a == b);
}

void main() {
  __vybeMain();
  __check('false\ntrue');
}
