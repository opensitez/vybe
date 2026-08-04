// vybe-test: dart/enum_enhanced/enhanced_enum_list_field_length
// origin: languages/dart/tests/dart/test_enum_enhanced.rs

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

enum Pack {
  small([1, 2]),
  big([1, 2, 3, 4]);
  final List<int> items;
  const Pack(this.items);
}
void __vybeMain() {
  __p(Pack.big.items.length);
}

void main() {
  __vybeMain();
  __check('4');
}
