// vybe-test: dart/factory_constructors_deep/factory_flyweight_reuses_equal_value_objects
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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

class SmallInt {
  static final Map<int, SmallInt> _pool = {};
  int value;
  SmallInt._(this.value);
  factory SmallInt(int v) {
    return _pool.putIfAbsent(v, () => SmallInt._(v));
  }
}
void __vybeMain() {
  __p(SmallInt(5) == SmallInt(5));
}

void main() {
  __vybeMain();
  __check('true');
}
