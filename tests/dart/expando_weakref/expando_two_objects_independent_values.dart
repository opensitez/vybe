// vybe-test: dart/expando_weakref/expando_two_objects_independent_values
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
  var a = Object();
  var b = Object();
  bag[a] = 10;
  bag[b] = 20;
  __p(bag[a]);
  __p(bag[b]);
}

void main() {
  __vybeMain();
  __check('10\n20');
}
