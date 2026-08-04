// vybe-test: dart/extension_types/extension_type_string_repeat_n_times
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

extension type Pattern(String value) {
  String repeat(int times) {
    return value * times;
  }
}
void __vybeMain() {
  Pattern p = Pattern('ab');
  __p(p.repeat(3));
}

void main() {
  __vybeMain();
  __check('ababab');
}
