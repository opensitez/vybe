// vybe-test: dart/cascades/cascade_custom_nested_list_field_mutated_via_chain
// origin: languages/dart/tests/dart/test_cascades.rs

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
  List<int> items = [];
  void seed(int n) { items.add(n); }
}
void __vybeMain() {
  var bag = Bag();
  bag..seed(1)..items.add(2)..items.add(3);
  __p(bag.items.join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
