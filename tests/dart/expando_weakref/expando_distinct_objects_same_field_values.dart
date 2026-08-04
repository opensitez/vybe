// vybe-test: dart/expando_weakref/expando_distinct_objects_same_field_values
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

class Pair {
  int a;
  int b;
  Pair(this.a, this.b);
}
void __vybeMain() {
  final bag = Expando<String>();
  var p1 = Pair(1, 2);
  var p2 = Pair(1, 2);
  bag[p1] = 'first';
  bag[p2] = 'second';
  __p(bag[p1]);
  __p(bag[p2]);
}

void main() {
  __vybeMain();
  __check('first\nsecond');
}
