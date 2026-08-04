// vybe-test: dart/expando_weakref/expando_function_object_as_key
// origin: languages/dart/tests/dart/test_expando_weakref.rs

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

void helper() {}
void __vybeMain() {
  final bag = Expando<String>();
  bag[helper] = 'fn';
  __p(bag[helper]);
}

void main() {
  __vybeMain();
  __check('fn');
}
