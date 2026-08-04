// vybe-test: dart/enum_enhanced/enhanced_enum_implements_two_methods
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

abstract class Named {
  String get label;
}
abstract class Valued {
  int get value;
}
enum Item implements Named, Valued {
  a(1),
  b(2);
  final int v;
  const Item(this.v);
  String get label => name;
  int get value => v;
}
void __vybeMain() {
  __p('${Item.b.label}:${Item.b.value}');
}

void main() {
  __vybeMain();
  __check('b:2');
}
