// vybe-test: dart/factory_constructors_deep/factory_creates_empty_via_private
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

class Bag {
  List<int> items;
  Bag._(this.items);
  factory Bag.empty() {
    return Bag._([]);
  }
}
void __vybeMain() {
  __p(Bag.empty().items.length);
}

void main() {
  __vybeMain();
  __check('0');
}
