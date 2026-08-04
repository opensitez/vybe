// vybe-test: dart/expando_weakref/expando_read_unrelated_object_returns_null
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
  var stored = Object();
  var other = Object();
  bag[stored] = 'data';
  __p(bag[other] == null);
}

void main() {
  __vybeMain();
  __check('true');
}
