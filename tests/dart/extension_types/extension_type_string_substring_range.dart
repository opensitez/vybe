// vybe-test: dart/extension_types/extension_type_string_substring_range
// origin: languages/dart/tests/dart/test_extension_types.rs

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

extension type Text(String value) {
  String slice(int start, int end) {
    return value.substring(start, end);
  }
}
void __vybeMain() {
  Text t = Text('hello');
  __p(t.slice(1, 4));
}

void main() {
  __vybeMain();
  __check('ell');
}
