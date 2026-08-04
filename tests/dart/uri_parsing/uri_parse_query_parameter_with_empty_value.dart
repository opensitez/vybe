// vybe-test: dart/uri_parsing/uri_parse_query_parameter_with_empty_value
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
  var u = Uri.parse('http://example.com/?flag=');
  __p(u.queryParameters['flag']);
  __p(u.queryParameters.containsKey('flag'));
}

void main() {
  __vybeMain();
  __check('\ntrue');
}
