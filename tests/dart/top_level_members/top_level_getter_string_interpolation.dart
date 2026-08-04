// vybe-test: dart/top_level_members/top_level_getter_string_interpolation
// origin: languages/dart/tests/dart/test_top_level_members.rs

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

String host = 'localhost';
String get endpoint {
  return 'http://$host';
}
void __vybeMain() {
  __p(endpoint.contains('localhost'));
}

void main() {
  __vybeMain();
  __check('true');
}
