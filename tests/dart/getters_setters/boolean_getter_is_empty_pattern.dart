// vybe-test: dart/getters_setters/boolean_getter_is_empty_pattern
// origin: languages/dart/tests/dart/test_getters_setters.rs

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
  List<int> _items = [];
  bool get isEmpty {
    return _items.isEmpty;
  }
}
void __vybeMain() {
  __p(Bag().isEmpty);
}

void main() {
  __vybeMain();
  __check('true');
}
