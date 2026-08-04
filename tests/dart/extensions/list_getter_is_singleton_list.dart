// vybe-test: dart/extensions/list_getter_is_singleton_list
// origin: languages/dart/tests/dart/test_extensions.rs

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

extension ListSingle on List<int> {
  bool get isSingleton => length == 1;
}
void __vybeMain() {
  __p([42].isSingleton);
}

void main() {
  __vybeMain();
  __check('true');
}
