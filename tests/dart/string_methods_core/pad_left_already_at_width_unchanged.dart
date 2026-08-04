// vybe-test: dart/string_methods_core/pad_left_already_at_width_unchanged
// origin: languages/dart/tests/dart/test_string_methods_core.rs

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
  __p('abc'.padLeft(3, '0'));
}

void main() {
  __vybeMain();
  __check('abc');
}
