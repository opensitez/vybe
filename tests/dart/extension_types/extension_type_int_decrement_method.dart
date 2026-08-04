// vybe-test: dart/extension_types/extension_type_int_decrement_method
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

extension type Counter(int count) {
  Counter decrement() {
    return Counter(count - 1);
  }
}
void __vybeMain() {
  Counter c = Counter(5);
  __p(c.decrement().count);
}

void main() {
  __vybeMain();
  __check('4');
}
