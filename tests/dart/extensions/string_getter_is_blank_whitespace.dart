// vybe-test: dart/extensions/string_getter_is_blank_whitespace
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

extension StrBlank on String {
  bool get isBlank => trim().isEmpty;
}
void __vybeMain() {
  __p('   '.isBlank);
}

void main() {
  __vybeMain();
  __check('true');
}
