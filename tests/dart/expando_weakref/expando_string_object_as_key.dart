// vybe-test: dart/expando_weakref/expando_string_object_as_key
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

void __vybeMain() {
  final bag = Expando<int>();
  var s = 'key';
  // Damaged test repaired: an Expando key CANNOT be a string — dart 3.10.4
  // throws "Invalid argument (object): Cannot be a string, number, boolean,
  // record, null, Pointer, Struct or Union" (measured). The original
  // expectation '7' never held under real dart.
  try {
    bag[s] = 7;
    __p(bag[s]);
  } catch (e) {
    __p('threw');
  }
}

void main() {
  __vybeMain();
  __check('threw');
}
