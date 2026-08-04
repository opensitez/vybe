// vybe-test: dart/type_casts/type_narrowing_is_not_string_routes_to_else
// origin: languages/dart/tests/dart/test_type_casts.rs

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

void describe(Object value) {
  if (value is! String) {
    __p('not-str');
  } else {
    __p('str');
  }
}
void __vybeMain() {
  describe(12);
}

void main() {
  __vybeMain();
  __check('not-str');
}
