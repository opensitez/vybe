// vybe-test: dart/getters_setters/setter_appends_to_private_list
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
  List<int> _items = [];
  set last(int v) {
    _items.add(v);
  }
  String dump() {
    return _items.join(',');
  }
}
void __vybeMain() {
  var c = Collection();
  c.last = 7;
  c.last = 8;
  __p(c.dump());
}

void main() {
  __vybeMain();
  __check('7,8');
}
