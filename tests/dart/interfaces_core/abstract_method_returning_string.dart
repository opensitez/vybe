// vybe-test: dart/interfaces_core/abstract_method_returning_string
// origin: languages/dart/tests/dart/test_interfaces_core.rs

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

abstract class Named {
  String name();
}
class Item implements Named {
  String name() {
    return 'item';
  }
}
void __vybeMain() {
  __p(Item().name());
}

void main() {
  __vybeMain();
  __check('item');
}
