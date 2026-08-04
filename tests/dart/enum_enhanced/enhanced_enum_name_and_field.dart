// vybe-test: dart/enum_enhanced/enhanced_enum_name_and_field
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

enum Fruit {
  apple(1),
  banana(2);
  final int count;
  const Fruit(this.count);
}
void __vybeMain() {
  var f = Fruit.banana;
  __p('${f.name}:${f.count}');
}

void main() {
  __vybeMain();
  __check('banana:2');
}
