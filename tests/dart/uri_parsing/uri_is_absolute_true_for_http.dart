// vybe-test: dart/uri_parsing/uri_is_absolute_true_for_http
// origin: languages/dart/tests/dart/test_uri_parsing.rs

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
  var u = Uri.parse('http://example.com/');
  __p(u.isAbsolute);
}

void main() {
  __vybeMain();
  __check('true');
}
