// vybe-test: dart/expando_weakref/expando_three_objects_three_entries
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
  var o1 = Object();
  var o2 = Object();
  var o3 = Object();
  bag[o1] = 1;
  bag[o2] = 2;
  bag[o3] = 3;
  __p(bag[o1] + bag[o2]! + bag[o3]!);
}

void main() {
  __vybeMain();
  __check('6');
}
