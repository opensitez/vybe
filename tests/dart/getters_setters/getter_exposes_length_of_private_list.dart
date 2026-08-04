// vybe-test: dart/getters_setters/getter_exposes_length_of_private_list
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

class Collection {
  List<int> _items = [1, 2, 3];
  int get length {
    return _items.length;
  }
}
void __vybeMain() {
  __p(Collection().length);
}

void main() {
  __vybeMain();
  __check('3');
}
