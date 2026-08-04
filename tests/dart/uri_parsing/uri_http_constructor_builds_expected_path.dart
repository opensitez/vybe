// vybe-test: dart/uri_parsing/uri_http_constructor_builds_expected_path
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
  var u = Uri.http('example.com', '/api');
  __p(u.scheme);
  __p(u.host);
  __p(u.path);
}

void main() {
  __vybeMain();
  __check('http\nexample.com\n/api');
}
