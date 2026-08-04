// vybe-test: dart/extensions/string_static_concat_two_strings
// origin: languages/dart/tests/dart/test_extensions.rs

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

extension StrJoin on String {
  static String concat(String a, String b) => a + b;
}
void __vybeMain() {
  __p(StrJoin.concat('foo', 'bar'));
}

void main() {
  __vybeMain();
  __check('foobar');
}
