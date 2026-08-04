// vybe-test: dart/expando_weakref/expando_null_key_on_set_throws
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
  final bag = Expando<String>();
  try {
    bag[null] = 'bad';
    __p('no throw');
  } catch (e) {
    __p('threw');
  }
}

void main() {
  __vybeMain();
  __check('threw');
}
