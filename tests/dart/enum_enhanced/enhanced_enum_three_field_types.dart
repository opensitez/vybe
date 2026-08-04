// vybe-test: dart/enum_enhanced/enhanced_enum_three_field_types
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

enum Record {
  entry(1, 'x', true);
  final int id;
  final String key;
  final bool active;
  const Record(this.id, this.key, this.active);
}
void __vybeMain() {
  __p('${Record.entry.id}:${Record.entry.key}:${Record.entry.active}');
}

void main() {
  __vybeMain();
  __check('1:x:true');
}
