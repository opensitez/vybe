// vybe-test: dart/generics_core/generic_function_list_repeat
// origin: languages/dart/tests/dart/test_generics_core.rs

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

List<T> repeat<T>(T value, int times) {
  return List.generate(times, (_) => value);
}
void __vybeMain() {
  var list = repeat(7, 3);
  __p(list.join(','));
}

void main() {
  __vybeMain();
  __check('7,7,7');
}
