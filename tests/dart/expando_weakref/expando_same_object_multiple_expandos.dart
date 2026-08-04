// vybe-test: dart/expando_weakref/expando_same_object_multiple_expandos
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
  final names = Expando<String>();
  final scores = Expando<int>();
  var obj = Object();
  names[obj] = 'alice';
  scores[obj] = 99;
  __p(names[obj]);
  __p(scores[obj]);
}

void main() {
  __vybeMain();
  __check('alice\n99');
}
