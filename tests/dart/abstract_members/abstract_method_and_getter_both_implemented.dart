// vybe-test: dart/abstract_members/abstract_method_and_getter_both_implemented
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Entity {
  String get id;
  String describe();
}
class Product extends Entity {
  String get id => 'p1';
  String describe() {
    return 'product';
  }
}
void __vybeMain() {
  __p(Product().describe());
}

void main() {
  __vybeMain();
  __check('product');
}
