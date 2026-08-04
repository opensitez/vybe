// vybe-test: dart/object_protocol/enum_identical_same_member
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

enum Status { ok, fail }
void __vybeMain() {
  __p(identical(Status.ok, Status.ok));
  __p(Status.ok == Status.ok);
}

void main() {
  __vybeMain();
  __check('true\ntrue');
}
